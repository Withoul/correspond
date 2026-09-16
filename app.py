import os
import sys
import shutil
import tempfile
import uuid
import json
import re
import html
import asyncio
import subprocess
import copy
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
import openpyxl
import mammoth
import zipfile
import win32com.client
import lxml.html
from jinja2 import Template
from docxtpl import DocxTemplate
from docx import Document
from docx.shared import Inches, Pt, RGBColor 
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import qn
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, BackgroundTasks
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

# Directorios de la aplicación
if getattr(sys, 'frozen', False):
    BASE_DIR = Path(sys._MEIPASS)
    EXE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).resolve().parent
    EXE_DIR = BASE_DIR

STATIC_DIR = BASE_DIR / "static"
TEMPLATES_DIR = BASE_DIR / "templates"
TEMP_DIR = EXE_DIR / "temp_uploads"
OUTPUT_BASE_DIR = EXE_DIR / "output"
PROJECTS_FILE = EXE_DIR / "projects_history.json"

for d in [STATIC_DIR, TEMPLATES_DIR, TEMP_DIR, OUTPUT_BASE_DIR]:
    d.mkdir(exist_ok=True)

# Limpiar archivos de temp_uploads en el arranque de la aplicación
for f in TEMP_DIR.glob("*"):
    if f.is_file():
        try:
            os.remove(f)
        except Exception:
            pass

app = FastAPI(title="Correspon - Automatización de Correspondencia")

# CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Montar estáticos
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

# Estado global de tareas en progreso
jobs_status = {}
# GESTIÓN DE AJUSTES EN PROYECTO (settings.json INDIVIDUAL POR CARPETA)
DEFAULT_PROJECT_SETTINGS = {
    "output_format": "docx",
    "output_filename_pattern": "{{ index }}_documento.docx",
    "outlook_enabled": False,
    "email_column": "Email",
    "check_column": "Enviado",
    "doc_check_column": "Doc_Generado",
    "doc_check_value_mode": "timestamp", # "timestamp", "custom_text", "format"
    "doc_check_custom_text": "Generado",
    "all_fields_mandatory": False,
    "mandatory_columns": [],
    "template_rules_enabled": False,
    "template_rules": [],
    "send_mode": "display",
    "attachment_type": "pdf",
    "email_subject": "Notificación Oficial - {{ Nombre }}",
    "email_body_source": "text",
    "email_body": "Estimado/a {{ Nombre }},\n\nAdjuntamos su documento oficial correspondiente a {{ Empresa }}.\n\nSaludos cordiales.",
    "custom_paragraph_rules": [],
    "additional_emails": [],
    "email_columns": [],
    "custom_menu_assignments": {"columns": {}, "cells": {}}
}

def load_project_settings(folder_path: str) -> dict:
    """Carga los ajustes desde <folder_path>/settings.json si existe."""
    try:
        sf = Path(folder_path) / "settings.json"
        if sf.exists():
            with open(sf, "r", encoding="utf-8") as f:
                data = json.load(f)
                res = DEFAULT_PROJECT_SETTINGS.copy()
                res.update(data)
                return res
    except Exception as e:
        print("Error leyendo settings.json:", e)
    return DEFAULT_PROJECT_SETTINGS.copy()

def save_project_settings(folder_path: str, settings: dict):
    """Guarda los ajustes en <folder_path>/settings.json."""
    try:
        sf = Path(folder_path) / "settings.json"
        sf.parent.mkdir(parents=True, exist_ok=True)
        with open(sf, "w", encoding="utf-8") as f:
            json.dump(settings, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print("Error guardando settings.json:", e)


def force_close_project_files(folder_path: str):
    """Fuerza el cierre de documentos de Word, Excel y PDF que pertenezcan a la carpeta especificada."""
    if sys.platform == "win32":
        folder_clean = os.path.normpath(folder_path).lower()
        try:
            import pythoncom
            pythoncom.CoInitialize()
            try:
                word = win32com.client.GetActiveObject("Word.Application")
                for doc in list(word.Documents):
                    if folder_clean in os.path.normpath(str(doc.FullName)).lower():
                        doc.Close(False)
            except Exception:
                pass
            try:
                excel = win32com.client.GetActiveObject("Excel.Application")
                for wb in list(excel.Workbooks):
                    if folder_clean in os.path.normpath(str(wb.FullName)).lower():
                        wb.Close(False)
            except Exception:
                pass
            pythoncom.CoUninitialize()
        except Exception as e:
            print("Error cerrando COM:", e)


# GESTIÓN DE HISTORIAL DE PROYECTOS / RUTAS MULTIPLES
def ensure_sample_project(p_data: dict) -> dict:
    """Garantiza que el proyecto demo de muestra exista con sus archivos Excel y Word válidos en disco."""
    projects = p_data.get("projects", [])
    
    projs_base_dir = EXE_DIR / "Proyectos_Correspondencia"
    projs_base_dir.mkdir(parents=True, exist_ok=True)

    # Mantener únicamente proyectos válidos cuyas carpetas existan en el disco
    valid_projects = [p for p in projects if p.get("folder_path") and os.path.exists(p.get("folder_path"))]

    demo_dir = projs_base_dir / "Correspondencia_Oficial_Demo"
    demo_dir.mkdir(parents=True, exist_ok=True)

    excel_path = demo_dir / "Datos_Ejemplo_Correspondencia.xlsx"
    docx_path = demo_dir / "Plantilla_Notificacion_Oficial.docx"
    docx2_path = demo_dir / "Plantilla_Certificado_Trabajo.docx"
    output_dir = demo_dir / "salida"
    output_dir.mkdir(exist_ok=True)

    # 1. Asegurar existencia del archivo Excel de muestra con 1 solo registro limpio
    df_sample = pd.DataFrame([
        {
            "Nombre": "Alexander Murillo",
            "Cargo": "Docente Investigador",
            "Empresa": "ITSJAPON",
            "Tipo_Documento": "Notificación",
            "Monto": "$2,500.00",
            "Fecha": "18/08/2026",
            "Email": "ajmurillol@itsjapon.edu.ec",
            "Doc_Generado": "",
            "Enviado": ""
        }
    ])
    df_sample.to_excel(excel_path, index=False)

    # 2. Asegurar existencia de las plantillas Word de muestra
    if not docx_path.exists():
        try:
            from docx import Document
            doc1 = Document()
            doc1.add_heading("NOTIFICACIÓN OFICIAL DE REQUERIMIENTO", level=1)
            doc1.add_paragraph("Fecha: {{ Fecha }}")
            doc1.add_paragraph("Estimado(a) {{ Nombre }},\n{{ Cargo }} - {{ Empresa }}")
            doc1.add_paragraph("Por medio de la presente, le notificamos oficialmente el estado de su trámite por un importe de {{ Monto }}.")
            doc1.add_paragraph("Quedamos a su disposición ante cualquier duda a través de su correo: {{ Email }}.")
            doc1.add_paragraph("Atentamente,\nDirección General")
            doc1.save(str(docx_path.resolve()))
        except Exception as ex1:
            print("Error creando plantilla 1:", ex1)

    if not docx2_path.exists():
        try:
            from docx import Document
            doc2 = Document()
            doc2.add_heading("CERTIFICADO DE TRABAJO Y DESEMPEÑO", level=1)
            doc2.add_paragraph("A QUIEN PUEDA INTERESAR:")
            doc2.add_paragraph("Se hace constar que el/la Sra.(a) {{ Nombre }}, se desempeña como {{ Cargo }} en la entidad {{ Empresa }}.")
            doc2.add_paragraph("Expedido a la fecha {{ Fecha }} para los fines que el interesado considere convenientes.")
            doc2.add_paragraph("Firma Autorizada:\nGerencia de Talento Humano")
            doc2.save(str(docx2_path.resolve()))
        except Exception as ex2:
            print("Error creando plantilla 2:", ex2)

    demo_settings = DEFAULT_PROJECT_SETTINGS.copy()
    demo_settings.update({
        "doc_check_column": "Doc_Generado",
        "doc_check_value_mode": "timestamp",
        "all_fields_mandatory": False,
        "mandatory_columns": ["Nombre", "Email"],
        "template_rules_enabled": True,
        "template_rules": [
            {
                "column": "Tipo_Documento",
                "value": "Certificado Trabajo",
                "template_filename": "Plantilla_Certificado_Trabajo.docx"
            }
        ]
    })
    save_project_settings(str(demo_dir.resolve()), demo_settings)

    active_docx_paths = [str(p.resolve()) for p in [docx_path, docx2_path] if p.exists()]

    demo_id = "proj_demo_full"
    demo_project_data = {
        "id": demo_id,
        "name": "Correspondencia Oficial Demo",
        "folder_path": str(demo_dir.resolve()),
        "excel_path": str(excel_path.resolve()) if excel_path.exists() else None,
        "docx_path": active_docx_paths[0] if active_docx_paths else None,
        "docx_paths": active_docx_paths,
        "output_path": str(output_dir.resolve()),
        "created_at": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"),
        "settings": demo_settings
    }

    # Encontrar si existe entrada previa del proyecto demo
    demo_idx = next((i for i, p in enumerate(valid_projects) if p.get("id") == demo_id or p.get("name") == "Correspondencia Oficial Demo"), -1)

    if demo_idx >= 0:
        # Reparar rutas en proyecto demo existente
        valid_projects[demo_idx]["excel_path"] = str(excel_path.resolve())
        valid_projects[demo_idx]["docx_path"] = active_docx_paths[0] if active_docx_paths else None
        valid_projects[demo_idx]["docx_paths"] = active_docx_paths
        valid_projects[demo_idx]["settings"] = demo_settings
    else:
        # Insertar proyecto demo
        valid_projects.insert(0, demo_project_data)

    seen_paths = set()
    deduped_projects = []
    for p in valid_projects:
        fp = p.get("folder_path")
        fp_norm = os.path.normpath(fp).lower() if fp else p.get("id")
        if fp_norm not in seen_paths:
            seen_paths.add(fp_norm)
            deduped_projects.append(p)

    p_data["projects"] = deduped_projects

    if not p_data.get("active_project_id") or not any(p["id"] == p_data.get("active_project_id") for p in deduped_projects):
        p_data["active_project_id"] = deduped_projects[0]["id"] if deduped_projects else None

    save_projects_data(p_data)
    return p_data


def load_projects_data() -> dict:
    if PROJECTS_FILE.exists():
        try:
            with open(PROJECTS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                projects = data.get("projects", [])
                valid_projects = [p for p in projects if p.get("folder_path") and os.path.exists(p.get("folder_path"))]
                data["projects"] = valid_projects
                if data.get("active_project_id") and not any(p["id"] == data.get("active_project_id") for p in valid_projects):
                    data["active_project_id"] = valid_projects[0]["id"] if valid_projects else None
                
                data = ensure_sample_project(data)
                return data
        except Exception:
            pass

    data = {"active_project_id": None, "projects": []}
    data = ensure_sample_project(data)
    return data

def save_projects_data(data: dict):
    try:
        with open(PROJECTS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print("Error al guardar proyectos_history.json:", e)


class GenerateRequest(BaseModel):
    excel_path: str
    docx_path: Optional[str] = None
    docx_paths: Optional[List[str]] = None
    output_filename_pattern: Optional[str] = "{{ index }}_documento.docx"
    output_dir_override: Optional[str] = None
    settings: Optional[dict] = None
    action: Optional[str] = "both" # "generate_docs", "send_emails", or "both"
    selected_row_indices: Optional[List[int]] = None
    generate_mode: Optional[str] = "all" # "all", "missing_only", or "selected"


class CancelJobRequest(BaseModel):
    job_id: str


class SaveExcelDataRequest(BaseModel):
    filepath: str
    columns: List[str]
    records: List[Dict[str, str]]


class SaveDocxContentRequest(BaseModel):
    docx_path: str
    content_html: Optional[str] = None
    html: Optional[str] = None


class CreateProjectRequest(BaseModel):
    name: str
    parent_dir: Optional[str] = None


class ImportProjectRequest(BaseModel):
    folder_path: str


class ListGeneratedFilesRequest(BaseModel):
    project_id: str


class UnlinkFileRequest(BaseModel):
    project_id: str
    docx_path: Optional[str] = None


class SelectProjectRequest(BaseModel):
    project_id: str


class SaveProjectSettingsRequest(BaseModel):
    project_id: str
    settings: dict


class DeleteProjectRequest(BaseModel):
    project_id: str


class DeleteTemplateRequest(BaseModel):
    project_id: str
    docx_path: str


class DeleteExcelRequest(BaseModel):
    project_id: str


@app.post("/api/delete-excel")
async def delete_excel(req: DeleteExcelRequest):
    """Elimina físicamente el archivo Excel asignado a un proyecto y actualiza el historial."""
    p_data = load_projects_data()
    selected = None
    for p in p_data.get("projects", []):
        if p["id"] == req.project_id:
            selected = p
            break

    if not selected:
        raise HTTPException(status_code=404, detail="Proyecto no encontrado.")

    excel_path_str = selected.get("excel_path")
    if excel_path_str and os.path.exists(excel_path_str):
        try:
            force_close_project_files(selected.get("folder_path", ""))
            os.remove(excel_path_str)
        except Exception:
            pass

    selected["excel_path"] = None
    save_projects_data(p_data)

    return {"success": True}


@app.post("/api/delete-project")
async def delete_project(req: DeleteProjectRequest):
    """Cierra todos los archivos abiertos (Word, Excel, PDF), elimina físicamente la carpeta del proyecto y actualiza el historial."""
    p_data = load_projects_data()
    projects = p_data.get("projects", [])
    target = None
    for p in projects:
        if p["id"] == req.project_id:
            target = p
            break

    if not target:
        raise HTTPException(status_code=404, detail="Proyecto no encontrado.")

    folder_path = target.get("folder_path")
    if folder_path and os.path.exists(folder_path):
        force_close_project_files(folder_path)
        import time
        time.sleep(0.3)
        try:
            shutil.rmtree(folder_path)
        except Exception:
            shutil.rmtree(folder_path, ignore_errors=True)

    # Limpiar cualquier vista previa PDF generada
    PDF_PREVIEW_CACHE.clear()
    for f in TEMP_DIR.glob("preview_*.pdf"):
        try:
            os.remove(f)
        except Exception:
            pass

    remaining = [p for p in projects if p["id"] != req.project_id]
    p_data["projects"] = remaining

    if p_data.get("active_project_id") == req.project_id:
        p_data["active_project_id"] = remaining[0]["id"] if remaining else None

    save_projects_data(p_data)

    return {
        "success": True,
        "active_project_id": p_data.get("active_project_id"),
        "projects": p_data.get("projects")
    }


@app.get("/", response_class=HTMLResponse)
async def read_root():
    index_path = TEMPLATES_DIR / "index.html"
    if not index_path.exists():
        return HTMLResponse(content="<h1>Cargando aplicación...</h1>", status_code=200)
    with open(index_path, "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())


@app.get("/api/projects")
async def get_projects():
    """Retorna la lista de proyectos/carpetas de correspondencia guardados."""
    data = load_projects_data()
    return data


@app.post("/api/save-project-settings")
async def save_project_settings_api(req: SaveProjectSettingsRequest):
    """Guarda la configuración y reglas de párrafos en el settings.json de la carpeta del proyecto."""
    p_data = load_projects_data()
    selected = None
    for p in p_data.get("projects", []):
        if p["id"] == req.project_id:
            selected = p
            break

    if not selected:
        raise HTTPException(status_code=404, detail="Proyecto no encontrado.")

    folder_path = selected.get("folder_path")
    if not folder_path:
        raise HTTPException(status_code=400, detail="El proyecto no tiene carpeta física asignada.")

    save_project_settings(folder_path, req.settings)
    selected["settings"] = req.settings
    save_projects_data(p_data)

    return {"success": True, "settings": req.settings}








@app.post("/api/create-project")
async def create_project(req: CreateProjectRequest):
    """Crea una nueva carpeta de proyecto EN BLANCO (sin Excel ni plantilla Word por defecto)."""
    clean_name = req.name.strip().replace("/", "_").replace("\\", "_")
    if not clean_name:
        raise HTTPException(status_code=400, detail="El nombre del proyecto no puede estar vacío.")

    parent_folder = Path(req.parent_dir) if req.parent_dir and os.path.exists(req.parent_dir) else (EXE_DIR / "Proyectos_Correspondencia")
    parent_folder.mkdir(exist_ok=True)

    project_dir = parent_folder / clean_name
    project_dir.mkdir(exist_ok=True)

    output_dir = project_dir / "salida"
    output_dir.mkdir(exist_ok=True)

    p_data = load_projects_data()
    proj_id = f"proj_{uuid.uuid4().hex[:8]}"

    initial_settings = DEFAULT_PROJECT_SETTINGS.copy()
    save_project_settings(str(project_dir.resolve()), initial_settings)

    new_project = {
        "id": proj_id,
        "name": clean_name,
        "folder_path": str(project_dir.resolve()),
        "excel_path": None,
        "docx_path": None,
        "docx_paths": [],
        "output_path": str(output_dir.resolve()),
        "created_at": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"),
        "settings": initial_settings
    }

    p_data["projects"].insert(0, new_project)
    p_data["active_project_id"] = proj_id
    save_projects_data(p_data)

    return {
        "success": True,
        "project": new_project,
        "columns": [],
        "records": [],
        "total_rows": 0,
        "html": "",
        "docx_templates": []
    }


@app.post("/api/import-project")
async def import_project(req: ImportProjectRequest):
    """Importa y registra una carpeta de proyecto existente en el disco sin borrar sus archivos."""
    raw_path = req.folder_path.strip().strip('"').strip("'")
    if not raw_path:
        raise HTTPException(status_code=400, detail="Debes ingresar la ruta de la carpeta del proyecto.")

    p_obj = Path(raw_path)
    if not p_obj.exists() or not p_obj.is_dir():
        raise HTTPException(status_code=400, detail="La carpeta especificada no existe en el disco.")

    folder_path_str = str(p_obj.resolve())
    p_data = load_projects_data()

    # Verificar si la carpeta ya está registrada
    existing = next((p for p in p_data.get("projects", []) if os.path.normpath(p.get("folder_path", "")).lower() == os.path.normpath(folder_path_str).lower()), None)
    if existing:
        p_data["active_project_id"] = existing["id"]
        save_projects_data(p_data)
        return {"success": True, "project": existing}

    found_excels = [f for f in p_obj.glob("*.xlsx") if not f.name.startswith("~$")] + [f for f in p_obj.glob("*.xls") if not f.name.startswith("~$")] + [f for f in p_obj.glob("*.xlsm") if not f.name.startswith("~$")]
    found_docxs = [f for f in p_obj.glob("*.docx") if not f.name.startswith("~$")]

    e_path = str(found_excels[0].resolve()) if found_excels else None
    d_paths = [str(f.resolve()) for f in found_docxs]
    out_path = str((p_obj / "salida").resolve())
    (p_obj / "salida").mkdir(exist_ok=True)

    st = load_project_settings(folder_path_str)
    proj_id = f"proj_{uuid.uuid4().hex[:8]}"

    new_proj_entry = {
        "id": proj_id,
        "name": p_obj.name.replace("_", " "),
        "folder_path": folder_path_str,
        "excel_path": e_path,
        "docx_path": d_paths[0] if d_paths else None,
        "docx_paths": d_paths,
        "output_path": out_path,
        "created_at": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"),
        "settings": st
    }

    p_data["projects"].insert(0, new_proj_entry)
    p_data["active_project_id"] = proj_id
    save_projects_data(p_data)

    return {"success": True, "project": new_proj_entry}


@app.post("/api/list-generated-files")
async def list_generated_files(req: ListGeneratedFilesRequest):
    """Retorna el listado visual de archivos Word y PDF generados en la carpeta salida/ del proyecto."""
    p_data = load_projects_data()
    selected = next((p for p in p_data.get("projects", []) if p["id"] == req.project_id), None)
    if not selected:
        raise HTTPException(status_code=404, detail="Proyecto no encontrado.")

    output_dir_str = selected.get("output_path")
    if not output_dir_str or not os.path.exists(output_dir_str):
        folder_obj = Path(selected.get("folder_path", "")) / "salida"
        folder_obj.mkdir(exist_ok=True)
        output_dir_str = str(folder_obj.resolve())

    output_dir = Path(output_dir_str)
    files = []

    if output_dir.exists():
        for f in output_dir.glob("**/*"):
            if f.is_file() and f.suffix.lower() in [".docx", ".pdf"] and not f.name.startswith("~$"):
                try:
                    stat = f.stat()
                    files.append({
                        "filename": f.name,
                        "filepath": str(f.resolve()),
                        "type": f.suffix.lower().replace(".", ""),
                        "size_kb": round(stat.st_size / 1024, 1),
                        "mtime": pd.Timestamp(stat.st_mtime, unit='s').strftime("%Y-%m-%d %H:%M:%S")
                    })
                except Exception:
                    pass

    files.sort(key=lambda x: x["mtime"], reverse=True)
    return {"success": True, "files": files, "output_dir": output_dir_str}


@app.post("/api/unlink-excel")
async def unlink_excel_api(req: UnlinkFileRequest):
    """Desvincula el archivo Excel del proyecto (sin eliminarlo del disco)."""
    p_data = load_projects_data()
    selected = next((p for p in p_data.get("projects", []) if p["id"] == req.project_id), None)
    if not selected:
        raise HTTPException(status_code=404, detail="Proyecto no encontrado.")

    selected["excel_path"] = None
    save_projects_data(p_data)
    return {"success": True}


@app.post("/api/unlink-docx")
async def unlink_docx_api(req: UnlinkFileRequest):
    """Desvincula una plantilla Word del proyecto (sin eliminarla del disco)."""
    p_data = load_projects_data()
    selected = next((p for p in p_data.get("projects", []) if p["id"] == req.project_id), None)
    if not selected:
        raise HTTPException(status_code=404, detail="Proyecto no encontrado.")

    paths = selected.get("docx_paths", [])
    if req.docx_path in paths:
        paths.remove(req.docx_path)
        selected["docx_paths"] = paths
        if selected.get("docx_path") == req.docx_path:
            selected["docx_path"] = paths[0] if paths else None
        save_projects_data(p_data)

    return {"success": True, "docx_paths": selected.get("docx_paths", [])}


class GetExcelMenusRequest(BaseModel):
    filepath: str


class SaveExcelMenusRequest(BaseModel):
    filepath: str
    menus: dict
    project_id: Optional[str] = None
    assignments: Optional[dict] = None


def get_excel_menus_from_file(filepath: Optional[str]) -> dict:
    """Lee todos los menús personalizados y sus opciones desde la hoja oculta menu_list_app."""
    if not filepath or not os.path.exists(filepath):
        return {}
    if not filepath.lower().endswith((".xlsx", ".xlsm", ".xls")):
        return {}
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        if "menu_list_app" not in wb.sheetnames:
            return {}
        ws = wb["menu_list_app"]
        menus = {}
        for col_idx in range(1, ws.max_column + 1):
            header_val = ws.cell(row=1, column=col_idx).value
            if header_val is None:
                continue
            menu_name = str(header_val).strip()
            if not menu_name:
                continue
            items = []
            for r in range(2, ws.max_row + 1):
                v = ws.cell(row=r, column=col_idx).value
                if v is not None:
                    s_val = str(v).strip()
                    if s_val:
                        items.append(s_val)
            menus[menu_name] = items
        return menus
    except Exception as e:
        print(f"Error leyendo menu_list_app de {filepath}: {e}")
        return {}


def save_excel_menus_to_file(filepath: str, menus: dict) -> bool:
    """Crea o actualiza la hoja oculta menu_list_app en el archivo Excel con los menús y opciones."""
    if not filepath or not os.path.exists(filepath):
        return False
    if not filepath.lower().endswith((".xlsx", ".xlsm", ".xls")):
        return False
    try:
        is_xlsm = filepath.lower().endswith(".xlsm")
        wb = openpyxl.load_workbook(filepath, keep_vba=is_xlsm)
        if "menu_list_app" in wb.sheetnames:
            ws = wb["menu_list_app"]
            for row in ws.iter_rows():
                for cell in row:
                    cell.value = None
        else:
            ws = wb.create_sheet(title="menu_list_app")

        # Asegurar que la hoja siempre esté oculta
        ws.sheet_state = "hidden"

        col_idx = 1
        for menu_name, options in menus.items():
            clean_name = str(menu_name).strip()
            if not clean_name:
                continue
            ws.cell(row=1, column=col_idx, value=clean_name)
            for opt_idx, opt in enumerate(options):
                opt_str = str(opt).strip()
                if opt_str:
                    ws.cell(row=opt_idx + 2, column=col_idx, value=opt_str)
            col_idx += 1

        wb.save(filepath)
        return True
    except PermissionError:
        raise HTTPException(
            status_code=409,
            detail="No se pudo guardar la hoja oculta menu_list_app porque el archivo Excel está abierto en Microsoft Excel. Por favor ciérralo e intenta nuevamente."
        )
    except Exception as e:
        print(f"Error guardando menu_list_app en {filepath}: {e}")
        raise HTTPException(status_code=500, detail=f"Error guardando menús en Excel: {str(e)}")


@app.post("/api/get-excel-menus")
async def api_get_excel_menus(req: GetExcelMenusRequest):
    """Retorna los menús y opciones almacenados en la hoja oculta menu_list_app del archivo Excel."""
    menus = get_excel_menus_from_file(req.filepath)
    return {"success": True, "menus": menus}


@app.post("/api/save-excel-menus")
async def api_save_excel_menus(req: SaveExcelMenusRequest):
    """Guarda los menús y opciones en la hoja oculta menu_list_app y sincroniza las asignaciones del proyecto."""
    save_excel_menus_to_file(req.filepath, req.menus)
    if req.project_id:
        p_data = load_projects_data()
        selected = next((p for p in p_data.get("projects", []) if p["id"] == req.project_id), None)
        if selected and selected.get("folder_path"):
            st = load_project_settings(selected["folder_path"])
            if req.assignments is not None:
                st["custom_menu_assignments"] = req.assignments
            st["custom_menus_cache"] = req.menus
            save_project_settings(selected["folder_path"], st)
            selected["settings"] = st
            save_projects_data(p_data)
    return {"success": True, "menus": req.menus}


@app.post("/api/reload-excel")
async def reload_excel(data: dict):
    """Vuelve a leer el archivo Excel/CSV desde el disco y retorna las columnas y registros actualizados."""
    filepath = data.get("filepath")
    if not filepath or not os.path.exists(filepath):
        raise HTTPException(status_code=400, detail="Archivo Excel no encontrado.")
    try:
        if filepath.lower().endswith((".xlsx", ".xls", ".xlsm")):
            df = pd.read_excel(filepath, engine='openpyxl' if filepath.lower().endswith(('.xlsx', '.xlsm')) else None)
        else:
            df = pd.read_csv(filepath)

        columns = [str(col).strip() for col in df.columns.tolist() if str(col).strip() and not str(col).startswith("Unnamed:")]
        records = df.fillna("").to_dict(orient="records")
        return {
            "success": True,
            "columns": columns,
            "total_rows": len(df),
            "records": records,
            "all_records": records,
            "menus": get_excel_menus_from_file(filepath)
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error leyendo archivo Excel: {str(e)}")


@app.post("/api/check-excel-status")
async def check_excel_status(data: dict):
    """Retorna la fecha de modificación (mtime) del archivo Excel para sincronización automática en tiempo real."""
    filepath = data.get("filepath")
    if not filepath or not os.path.exists(filepath):
        return {"exists": False, "mtime": 0}
    try:
        mtime = os.path.getmtime(filepath)
        return {"exists": True, "mtime": mtime}
    except Exception:
        return {"exists": False, "mtime": 0}


def docx_to_editor_html(docx_path: str) -> str:
    """Extrae el HTML de un documento Word (.docx) preservando párrafos vacíos y enriqueciendo alineaciones y estilos para el editor."""
    if not docx_path or not os.path.exists(docx_path):
        return ""
    try:
        with open(docx_path, "rb") as f:
            res = mammoth.convert_to_html(f, ignore_empty_paragraphs=False)
            html_raw = res.value
    except Exception as ex:
        print(f"Error en extracción mammoth para {docx_path}: {ex}")
        return ""

    try:
        doc = Document(docx_path)
        root = lxml.html.fragment_fromstring(f"<div>{html_raw}</div>")
        p_idx = 0
        for elem in root:
            tag = elem.tag.lower() if isinstance(elem.tag, str) else ''
            if tag in ['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'blockquote']:
                if p_idx < len(doc.paragraphs):
                    dp = doc.paragraphs[p_idx]
                    p_idx += 1
                    style_parts = []
                    if dp.alignment == WD_ALIGN_PARAGRAPH.CENTER:
                        style_parts.append('text-align: center;')
                    elif dp.alignment == WD_ALIGN_PARAGRAPH.RIGHT:
                        style_parts.append('text-align: right;')
                    elif dp.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
                        style_parts.append('text-align: justify;')
                    
                    if style_parts:
                        existing_style = elem.get('style', '')
                        elem.set('style', (existing_style + ' ' + ' '.join(style_parts)).strip())

        return ''.join(lxml.html.tostring(child, encoding='unicode') for child in root)
    except Exception as e:
        print(f"Error enriqueciendo HTML para {docx_path}: {e}")
        return html_raw


@app.post("/api/select-project")
async def select_project(req: SelectProjectRequest):
    """Selecciona un proyecto existente y carga sus datos, plantillas Word y settings.json individual."""
    p_data = load_projects_data()
    selected = None
    for p in p_data.get("projects", []):
        if p["id"] == req.project_id:
            selected = p
            break

    if not selected:
        raise HTTPException(status_code=404, detail="Proyecto no encontrado.")

    folder_path_obj = Path(selected.get("folder_path", ""))

    # Auto-descubrimiento de Excel en carpeta si excel_path es invalido
    excel_path_str = selected.get("excel_path")
    excel_path = Path(excel_path_str) if excel_path_str else None
    
    if (not excel_path or not excel_path.exists()) and folder_path_obj.exists():
        found_excels = [f for f in folder_path_obj.glob("*.xlsx") if not f.name.startswith("~$")] + [f for f in folder_path_obj.glob("*.xlsm") if not f.name.startswith("~$")]
        if found_excels:
            excel_path = found_excels[0]
            selected["excel_path"] = str(excel_path.resolve())
            save_projects_data(p_data)

    columns = []
    records = []
    total_rows = 0

    if excel_path and excel_path.exists():
        try:
            if str(excel_path).lower().endswith((".xlsx", ".xls", ".xlsm")):
                df = pd.read_excel(excel_path, engine='openpyxl' if str(excel_path).lower().endswith(('.xlsx', '.xlsm')) else None)
            else:
                df = pd.read_csv(excel_path)
            columns = [str(col).strip() for col in df.columns.tolist() if str(col).strip() and not str(col).startswith("Unnamed:")]
            records = df.fillna("").to_dict(orient="records")
            total_rows = len(df)
        except Exception as ex_read:
            print("Error leyendo Excel en select_project:", ex_read)

    raw_docx_paths = selected.get("docx_paths") or ([selected["docx_path"]] if selected.get("docx_path") else [])
    valid_docx_paths = [p for p in raw_docx_paths if os.path.exists(p)]

    # Auto-descubrimiento de Plantillas Word en la carpeta del proyecto
    if folder_path_obj.exists():
        found_docx = [str(f.resolve()) for f in folder_path_obj.glob("*.docx") if not f.name.startswith("~$") and f.parent == folder_path_obj]
        for fd in found_docx:
            if fd not in valid_docx_paths:
                valid_docx_paths.append(fd)
        if valid_docx_paths:
            selected["docx_path"] = valid_docx_paths[0]
            selected["docx_paths"] = valid_docx_paths
        else:
            selected["docx_path"] = None
            selected["docx_paths"] = []
        save_projects_data(p_data)

    docx_templates = []
    for dp in valid_docx_paths:
        p_obj = Path(dp)
        try:
            html_content = docx_to_editor_html(str(p_obj.resolve()))
            docx_templates.append({
                "filepath": str(p_obj.resolve()),
                "filename": p_obj.name,
                "html": html_content
            })
        except Exception as ex_mam:
            print("Error leyendo plantilla:", ex_mam)

    first_html = docx_templates[0]["html"] if docx_templates else ""

    # Cargar los ajustes desde el settings.json de la carpeta propia del proyecto
    folder_path = selected.get("folder_path")
    if folder_path:
        project_settings = load_project_settings(folder_path)
        selected["settings"] = project_settings

    p_data["active_project_id"] = req.project_id
    save_projects_data(p_data)

    excel_p = selected.get("excel_path")
    menus = get_excel_menus_from_file(excel_p) if excel_p else {}

    return {
        "success": True,
        "project": selected,
        "columns": columns,
        "records": records,
        "total_rows": total_rows,
        "html": first_html,
        "docx_templates": docx_templates,
        "menus": menus
    }




PDF_PREVIEW_CACHE = {}

@app.post("/api/get-pdf-preview")
async def get_pdf_preview(data: dict):
    """Convierte la plantilla .docx actual a vista previa PDF estándar con caché por mtime."""
    docx_path_str = data.get("docx_path")
    if not docx_path_str or not os.path.exists(docx_path_str):
        raise HTTPException(status_code=400, detail="Documento DOCX no encontrado.")

    docx_path = Path(docx_path_str).resolve()
    resolved_key = str(docx_path)
    current_mtime = os.path.getmtime(docx_path)
    force_refresh = data.get("force_refresh", False)

    # Verificar si ya existe en caché y no se ha modificado la plantilla Word
    if not force_refresh and resolved_key in PDF_PREVIEW_CACHE:
        cached = PDF_PREVIEW_CACHE[resolved_key]
        if cached.get("mtime") == current_mtime:
            cache_pdf_file = TEMP_DIR / cached["filename"]
            if cache_pdf_file.exists():
                return {
                    "success": True,
                    "pdf_url": cached["pdf_url"],
                    "filename": cached["filename"],
                    "cached": True
                }
    # Eliminar el archivo de previsualización anterior si existe en caché antes de generar uno nuevo
    if resolved_key in PDF_PREVIEW_CACHE:
        old_cached = PDF_PREVIEW_CACHE[resolved_key]
        old_pdf_file = TEMP_DIR / old_cached["filename"]
        if old_pdf_file.exists():
            try:
                os.remove(old_pdf_file)
            except Exception as e:
                print(f"Error removiendo archivo de previsualizacion antiguo {old_pdf_file}: {e}")

    pdf_filename = f"preview_{uuid.uuid4().hex[:8]}.pdf"
    pdf_path = TEMP_DIR / pdf_filename

    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        word = None
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            wdoc = word.Documents.Open(str(docx_path.resolve()), ReadOnly=True, ConfirmConversions=False)
            wdoc.SaveAs2(str(pdf_path.resolve()), FileFormat=17) # 17 = wdFormatPDF
            wdoc.Close(0) # 0 = wdDoNotSaveChanges
        except Exception as ex_com:
            print("Error en Word COM para preview PDF, intentando docx2pdf fallback:", ex_com)
            try:
                from docx2pdf import convert
                convert(str(docx_path.resolve()), str(pdf_path.resolve()))
            except Exception as ex_fallback:
                raise Exception(f"No se pudo convertir Word a PDF: {str(ex_com)} / Fallback: {str(ex_fallback)}")
        finally:
            if word:
                try:
                    word.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()

        if not pdf_path.exists():
            raise Exception("No se generó el archivo PDF final.")

        res_data = {
            "success": True,
            "pdf_url": f"/api/view-pdf/{pdf_filename}",
            "filename": pdf_filename,
            "cached": False
        }

        # Guardar en el diccionario global de caché
        PDF_PREVIEW_CACHE[resolved_key] = {
            "mtime": current_mtime,
            "pdf_url": res_data["pdf_url"],
            "filename": pdf_filename
        }

        return res_data
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error generando vista previa PDF: {str(e)}")


@app.get("/api/view-pdf/{filename}")
async def view_pdf(filename: str):
    pdf_file = TEMP_DIR / filename
    if pdf_file.exists():
        return FileResponse(path=str(pdf_file), media_type="application/pdf")
    raise HTTPException(status_code=404, detail="Archivo PDF no encontrado.")


@app.post("/api/upload-excel")
async def upload_excel(file: UploadFile = File(...), project_id: Optional[str] = Form(None)):
    """Recibe un archivo .xlsx, .xlsm, .xls o .csv y retorna sus columnas y registros."""
    fn_lower = file.filename.lower()
    if not fn_lower.endswith((".xlsx", ".xls", ".csv", ".xlsm")):
        raise HTTPException(status_code=400, detail="Formato no soportado. Debe ser .xlsx, .xlsm, .xls o .csv")

    p_data = load_projects_data()
    selected_proj = None
    if project_id:
        for p in p_data.get("projects", []):
            if p["id"] == project_id:
                selected_proj = p
                break

    if selected_proj and selected_proj.get("folder_path"):
        proj_dir = Path(selected_proj["folder_path"])
        saved_path = proj_dir / file.filename
    else:
        temp_file_id = f"excel_{uuid.uuid4().hex[:8]}_{file.filename}"
        saved_path = TEMP_DIR / temp_file_id

    with open(saved_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    try:
        if fn_lower.endswith((".xlsx", ".xls", ".xlsm")):
            df = pd.read_excel(saved_path, engine='openpyxl' if fn_lower.endswith(('.xlsx', '.xlsm')) else None)
        else:
            df = pd.read_csv(saved_path)

        columns = [str(col).strip() for col in df.columns.tolist() if str(col).strip() and not str(col).startswith("Unnamed:")]
        if not columns:
            columns = ["Nombre", "Detalle", "Enviado"]
            df = pd.DataFrame([{"Nombre": "", "Detalle": "", "Enviado": "NO"}])

        total_rows = len(df)
        preview_data = df.head(5).fillna("").to_dict(orient="records")
        all_records = df.fillna("").to_dict(orient="records")

        if selected_proj:
            selected_proj["excel_path"] = str(saved_path.resolve())
            save_projects_data(p_data)

        return {
            "success": True,
            "filename": file.filename,
            "filepath": str(saved_path.resolve()),
            "columns": columns,
            "total_rows": total_rows,
            "preview": preview_data,
            "all_records": all_records,
            "menus": get_excel_menus_from_file(str(saved_path.resolve()))
        }

    except Exception as e:
        if saved_path.exists():
            os.remove(saved_path)
        raise HTTPException(status_code=500, detail=f"Error leyendo archivo Excel: {str(e)}")


def ensure_valid_docx(file_path: Path) -> Path:
    """Valida si un archivo es un .docx ZIP estructurado. Si está vacío o es un .doc antiguo, lo inicializa/convierte a .docx válido."""
    if file_path.stat().st_size == 0:
        doc = Document()
        doc.add_paragraph("")
        doc.save(str(file_path.resolve()))
        return file_path

    if zipfile.is_zipfile(file_path):
        try:
            with zipfile.ZipFile(file_path, 'r') as z:
                if 'word/document.xml' in z.namelist():
                    return file_path
        except Exception:
            pass

    converted_path = file_path.with_name(f"converted_{file_path.name}.docx")
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        word = win32com.client.DispatchEx("Word.Application")
        doc = word.Documents.Open(str(file_path.resolve()))
        doc.SaveAs2(str(converted_path.resolve()), FileFormat=16)
        doc.Close()
        word.Quit()
        pythoncom.CoUninitialize()
        return converted_path
    except Exception:
        try:
            doc = Document()
            doc.add_paragraph("")
            doc.save(str(converted_path.resolve()))
            return converted_path
        except Exception:
            if converted_path.exists():
                os.remove(converted_path)
            raise ValueError("El archivo subido no es un documento Word válido.")


@app.post("/api/upload-docx")
async def upload_docx(file: UploadFile = File(...)):
    """Recibe un archivo .docx o .doc, valida su estructura ZIP y extrae su HTML via Mammoth."""
    filename_lower = file.filename.lower()
    if not (filename_lower.endswith(".docx") or filename_lower.endswith(".doc")):
        raise HTTPException(status_code=400, detail="Formato no soportado. Debe ser .docx o .doc")

    temp_file_id = f"docx_{uuid.uuid4().hex[:8]}_{file.filename}"
    saved_path = TEMP_DIR / temp_file_id

    with open(saved_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    if saved_path.stat().st_size == 0:
        doc = Document()
        doc.add_paragraph("")
        doc.save(str(saved_path.resolve()))

    try:
        valid_docx_path = ensure_valid_docx(saved_path)
        html_content = docx_to_editor_html(str(valid_docx_path.resolve()))

        return {
            "success": True,
            "filename": file.filename,
            "filepath": str(valid_docx_path.resolve()),
            "html": html_content,
            "warnings": [],
        }
    except ValueError as ve:
        if saved_path.exists():
            os.remove(saved_path)
        raise HTTPException(status_code=400, detail=str(ve))
    except Exception as e:
        if saved_path.exists():
            os.remove(saved_path)
        raise HTTPException(status_code=500, detail=f"Error al procesar la plantilla Word: {str(e)}")


def update_excel_preserving_styles(filepath: str, columns: List[str], records: List[Dict[str, str]]):
    """Actualiza las celdas del archivo Excel (.xlsx o .xlsm) preservando 100% fuentes, bordes, colores, alineación y macros (VBA)."""
    filepath_lower = filepath.lower()
    if filepath_lower.endswith((".xlsx", ".xlsm", ".xls")):
        is_xlsm = filepath_lower.endswith(".xlsm")
        wb = openpyxl.load_workbook(filepath, keep_vba=is_xlsm)
        
        # Seleccionar la primera hoja de datos (que no sea la hoja oculta menu_list_app)
        sheet = None
        for s in wb.worksheets:
            if s.title != "menu_list_app":
                sheet = s
                break
        if sheet is None:
            sheet = wb.active

        # 1. Encabezados (Fila 1)
        for c_idx, col_name in enumerate(columns):
            cell = sheet.cell(row=1, column=c_idx + 1)
            cell.value = col_name

        max_existing_rows = sheet.max_row
        target_rows_count = len(records)

        # 2. Datos (Filas 2 en adelante)
        for r_idx, rec in enumerate(records):
            target_row = r_idx + 2
            for c_idx, col_name in enumerate(columns):
                val = rec.get(col_name, "")
                cell = sheet.cell(row=target_row, column=c_idx + 1)
                
                # Si la fila es nueva y supera la cantidad original, copiar el estilo de la fila previa
                if target_row > max_existing_rows and target_row > 2:
                    ref_cell = sheet.cell(row=target_row - 1, column=c_idx + 1)
                    if ref_cell.has_style:
                        cell.font = ref_cell.font.copy()
                        cell.fill = ref_cell.fill.copy()
                        cell.border = ref_cell.border.copy()
                        cell.alignment = ref_cell.alignment.copy()
                        cell.number_format = ref_cell.number_format
                
                cell.value = val

        # 3. Limpiar celdas sobrantes si la nueva tabla tiene menos filas que la original
        if max_existing_rows > target_rows_count + 1:
            for extra_row in range(target_rows_count + 2, max_existing_rows + 1):
                for c_idx in range(1, max(len(columns) + 1, sheet.max_column + 1)):
                    sheet.cell(row=extra_row, column=c_idx).value = None

        if "menu_list_app" in wb.sheetnames:
            wb["menu_list_app"].sheet_state = "hidden"

        wb.save(filepath)
    else:
        df = pd.DataFrame(records)
        for col in columns:
            if col not in df.columns:
                df[col] = ""
        if columns:
            df = df[columns]
        df.to_csv(filepath, index=False)


@app.post("/api/update-excel-data")
async def update_excel_data(req: SaveExcelDataRequest):
    """Guarda las modificaciones realizadas en la tabla Excel sin alterar estructura, fuentes ni colores."""
    try:
        update_excel_preserving_styles(req.filepath, req.columns, req.records)

        if req.filepath.lower().endswith((".xlsx", ".xls", ".xlsm")):
            df = pd.read_excel(req.filepath, engine='openpyxl' if req.filepath.lower().endswith(('.xlsx', '.xlsm')) else None)
        else:
            df = pd.read_csv(req.filepath)

        return {
            "success": True,
            "columns": req.columns,
            "total_rows": len(df),
            "preview": df.head(5).fillna("").to_dict(orient="records"),
            "all_records": df.fillna("").to_dict(orient="records"),
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error guardando datos Excel: {str(e)}")

def parse_style_attribute(style_str: str) -> dict:
    """Parsea una cadena CSS de atributo style en un diccionario clave-valor."""
    props = {}
    if not style_str:
        return props
    for item in style_str.split(';'):
        if ':' in item:
            k, v = item.split(':', 1)
            props[k.strip().lower()] = v.strip()
    return props


def get_alignment_from_html(elem):
    """Detecta alineación de párrafo a partir de atributos HTML o estilos CSS."""
    style = elem.get('style', '').lower() if hasattr(elem, 'get') else ''
    align_attr = elem.get('align', '').lower() if hasattr(elem, 'get') else ''
    if 'text-align: center' in style or align_attr == 'center':
        return WD_ALIGN_PARAGRAPH.CENTER
    elif 'text-align: right' in style or align_attr == 'right':
        return WD_ALIGN_PARAGRAPH.RIGHT
    elif 'text-align: justify' in style or align_attr == 'justify':
        return WD_ALIGN_PARAGRAPH.JUSTIFY
    elif 'text-align: left' in style or align_attr == 'left':
        return WD_ALIGN_PARAGRAPH.LEFT
    return None


def extract_segments_from_html_elem(elem) -> list:
    """Extrae recursivamente segmentos de texto con sus propiedades de formato (negrita, cursiva, color, fuente, saltos)."""
    segments = []

    def recurse(node, inherited):
        tag = node.tag.lower() if isinstance(node.tag, str) else ''
        current = copy.deepcopy(inherited)
        
        style = parse_style_attribute(node.get('style', '')) if hasattr(node, 'get') else {}
        
        if tag in ['strong', 'b'] or 'bold' in style.get('font-weight', '').lower() or style.get('font-weight', '') in ['700', '800', '900']:
            current['bold'] = True
        if tag in ['em', 'i'] or 'italic' in style.get('font-style', '').lower():
            current['italic'] = True
        if tag in ['u', 'ins'] or 'underline' in style.get('text-decoration', '').lower():
            current['underline'] = True
        if tag in ['s', 'strike', 'del'] or 'line-through' in style.get('text-decoration', '').lower():
            current['strike'] = True
            
        color_val = style.get('color', '')
        if color_val:
            m_hex = re.search(r'#([0-9a-fA-F]{6})', color_val)
            if m_hex:
                current['color'] = m_hex.group(1).upper()
            else:
                m_rgb = re.search(r'rgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)', color_val)
                if m_rgb:
                    r, g, b = int(m_rgb.group(1)), int(m_rgb.group(2)), int(m_rgb.group(3))
                    current['color'] = f'{r:02X}{g:02X}{b:02X}'

        font_fam = style.get('font-family', '')
        if font_fam:
            current['font_name'] = font_fam.split(',')[0].strip('\'" ')

        if tag == 'br':
            segments.append({'text': '', 'is_break': True, **current})
            
        if node.text:
            text = html.unescape(node.text)
            if text:
                segments.append({'text': text, 'is_break': False, **current})
                
        for child in node:
            recurse(child, current)
            if child.tail:
                tail_text = html.unescape(child.tail)
                if tail_text:
                    segments.append({'text': tail_text, 'is_break': False, **inherited})

    base_props = {'bold': None, 'italic': None, 'underline': None, 'strike': None, 'color': None, 'font_name': None}
    recurse(elem, base_props)
    return segments


def extract_original_run_styles(paragraph) -> tuple:
    """Extrae perfiles de formato de los runs existentes en el párrafo para permitir herencia exacta."""
    profiles = []
    default_profile = {
        'font_name': None,
        'font_size': None,
        'color': None,
        'bold': None,
        'italic': None,
        'underline': None,
        'strike': None,
        'rPr_xml': None
    }
    
    for r in paragraph.runs:
        color_hex = None
        if r.font.color and r.font.color.rgb:
            color_hex = str(r.font.color.rgb).upper()
        elif r._r.rPr is not None:
            c_elem = r._r.rPr.find(qn('w:color'))
            if c_elem is not None and c_elem.get(qn('w:val')):
                color_hex = c_elem.get(qn('w:val')).upper()

        prof = {
            'text': r.text,
            'font_name': r.font.name,
            'font_size': r.font.size,
            'color': color_hex,
            'bold': r.bold,
            'italic': r.italic,
            'underline': r.underline,
            'strike': r.font.strike,
            'rPr_xml': r._r.rPr.xml if r._r.rPr is not None else None
        }
        profiles.append(prof)
        if not default_profile['font_name'] and r.font.name:
            default_profile['font_name'] = r.font.name
        if not default_profile['font_size'] and r.font.size:
            default_profile['font_size'] = r.font.size
        if not default_profile['color'] and color_hex:
            default_profile['color'] = color_hex
            
    return profiles, default_profile


def update_paragraph_with_segments(paragraph, segments: list, explicit_align=None):
    """Actualiza los runs de un párrafo DOCX preservando el formato original y heredando tipografía/color de palabras adyacentes."""
    if explicit_align is not None:
        paragraph.alignment = explicit_align

    orig_profiles, default_prof = extract_original_run_styles(paragraph)
    
    # Limpiar runs existentes del párrafo sin tocar pPr (alineación, sangrías, interlineado, numeración)
    p_elem = paragraph._p
    for r_elem in list(p_elem.xpath('./w:r')):
        p_elem.remove(r_elem)
        
    if not segments:
        return

    prev_style = default_prof.copy()
    
    for idx, seg in enumerate(segments):
        if seg.get('is_break'):
            r = paragraph.add_run()
            r.add_break()
            continue
            
        seg_text = seg.get('text', '')
        if not seg_text:
            continue
            
        # Determinar el perfil de estilo base:
        # 1. ¿Coincide exactamente con el texto de un run original?
        matched_prof = None
        for op in orig_profiles:
            if op['text'] and op['text'].strip() == seg_text.strip():
                matched_prof = op
                break
                
        # 2. Si no coincide exactamente, verificar si existe un perfil en la misma posición ordinal
        if not matched_prof and idx < len(orig_profiles):
            matched_prof = orig_profiles[idx]
            
        # 3. Si no, heredar del run previo adyacente ("como escribir por encima")
        base_style = matched_prof if matched_prof else prev_style
        
        run = paragraph.add_run(seg_text)
        
        # Aplicar fuente y tamaño heredados
        if base_style.get('font_name'):
            run.font.name = base_style['font_name']
        if base_style.get('font_size'):
            run.font.size = base_style['font_size']
            
        # Aplicar color: explícito del HTML o heredado
        final_color = seg.get('color') or base_style.get('color')
        if final_color:
            try:
                run.font.color.rgb = RGBColor.from_string(final_color)
            except Exception:
                pass
                
        # Aplicar negrita: explícito del HTML o heredado
        if seg.get('bold') is not None:
            run.bold = seg['bold']
        elif base_style.get('bold') is not None:
            run.bold = base_style['bold']
            
        # Aplicar cursiva
        if seg.get('italic') is not None:
            run.italic = seg['italic']
        elif base_style.get('italic') is not None:
            run.italic = base_style['italic']

        # Aplicar subrayado
        if seg.get('underline') is not None:
            run.underline = seg['underline']
        elif base_style.get('underline') is not None:
            run.underline = base_style['underline']

        # Aplicar tachado
        if seg.get('strike') is not None:
            run.font.strike = seg['strike']
        elif base_style.get('strike') is not None:
            run.font.strike = base_style['strike']

        # Actualizar prev_style para el siguiente segmento adyacente
        prev_style = {
            'font_name': run.font.name or base_style.get('font_name'),
            'font_size': run.font.size or base_style.get('font_size'),
            'color': final_color,
            'bold': run.bold,
            'italic': run.italic,
            'underline': run.underline,
            'strike': run.font.strike,
            'rPr_xml': None
        }


def sync_html_to_docx(html_content: str, docx_path: str):
    """Actualiza el archivo DOCX existente in-place aplicando los cambios de texto/formato sin destruir la plantilla."""
    if os.path.exists(docx_path):
        doc = Document(docx_path)
    else:
        doc = Document()

    if not html_content or not html_content.strip():
        for p in doc.paragraphs:
            p.text = ''
        doc.save(docx_path)
        return

    raw_html = html_content.strip()
    try:
        root = lxml.html.fragment_fromstring(f"<div>{raw_html}</div>")
    except Exception:
        try:
            root = lxml.html.fromstring(f"<html><body><div>{raw_html}</div></body></html>")
            root = root.find(".//div") or root
        except Exception:
            clean_lines = re.split(r'</p>|</div>|<br\s*/?>', raw_html, flags=re.IGNORECASE)
            for i, line in enumerate(clean_lines):
                clean_text = html.unescape(re.sub(r'<[^>]+>', '', line)).strip()
                if i < len(doc.paragraphs):
                    doc.paragraphs[i].text = clean_text
                else:
                    doc.add_paragraph(clean_text)
            doc.save(docx_path)
            return

    # Extraer bloques de HTML (párrafos, encabezados, tablas, listas)
    html_blocks = []
    for child in root:
        tag = child.tag.lower() if isinstance(child.tag, str) else ''
        if tag in ['ul', 'ol']:
            for li in child.iterchildren():
                if isinstance(li.tag, str) and li.tag.lower() == 'li':
                    html_blocks.append(('li', li))
        elif tag == 'table':
            html_blocks.append(('table', child))
        else:
            html_blocks.append(('p', child))

    docx_p_list = doc.paragraphs
    p_cursor = 0
    t_cursor = 0

    for block_type, html_elem in html_blocks:
        if block_type in ['p', 'li']:
            segs = extract_segments_from_html_elem(html_elem)
            explicit_align = get_alignment_from_html(html_elem)
            
            if p_cursor < len(docx_p_list):
                target_p = docx_p_list[p_cursor]
                update_paragraph_with_segments(target_p, segs, explicit_align)
                p_cursor += 1
            else:
                new_p = doc.add_paragraph()
                update_paragraph_with_segments(new_p, segs, explicit_align)
                docx_p_list = doc.paragraphs
                p_cursor += 1
                
        elif block_type == 'table':
            rows_html = html_elem.xpath('.//tr')
            if t_cursor < len(doc.tables):
                target_tbl = doc.tables[t_cursor]
                t_cursor += 1
                
                for r_idx, r_elem in enumerate(rows_html):
                    cells_html = r_elem.xpath('./td | ./th')
                    if r_idx < len(target_tbl.rows):
                        row_obj = target_tbl.rows[r_idx]
                    else:
                        row_obj = target_tbl.add_row()
                        
                    for c_idx, c_elem in enumerate(cells_html):
                        if c_idx < len(row_obj.cells):
                            cell_obj = row_obj.cells[c_idx]
                            cell_segs = extract_segments_from_html_elem(c_elem)
                            cell_align = get_alignment_from_html(c_elem)
                            if cell_obj.paragraphs:
                                update_paragraph_with_segments(cell_obj.paragraphs[0], cell_segs, cell_align)
                            else:
                                cp = cell_obj.add_paragraph()
                                update_paragraph_with_segments(cp, cell_segs, cell_align)
            else:
                num_cols = max(len(r.xpath('./td | ./th')) for r in rows_html) if rows_html else 1
                new_tbl = doc.add_table(rows=0, cols=num_cols)
                new_tbl.style = 'Table Grid'
                for r_elem in rows_html:
                    cells_html = r_elem.xpath('./td | ./th')
                    row_cells = new_tbl.add_row().cells
                    for c_idx, c_elem in enumerate(cells_html):
                        if c_idx < len(row_cells):
                            cell_segs = extract_segments_from_html_elem(c_elem)
                            cell_align = get_alignment_from_html(c_elem)
                            update_paragraph_with_segments(row_cells[c_idx].paragraphs[0], cell_segs, cell_align)
                t_cursor += 1

    # Limpiar párrafos sobrantes si el HTML tiene menos párrafos
    while p_cursor < len(docx_p_list):
        docx_p_list[p_cursor].text = ''
        p_cursor += 1

    doc.save(docx_path)


# Alias de compatibilidad
convert_html_to_docx = sync_html_to_docx


@app.post("/api/save-docx-content")
async def save_docx_content(req: SaveDocxContentRequest):
    """Guarda las modificaciones realizadas en el editor A4 in-place en el archivo .docx maestro preservando todo el formato original."""
    docx_path = req.docx_path
    if not docx_path or not os.path.exists(docx_path):
        raise HTTPException(status_code=400, detail="El archivo plantilla DOCX especificado no existe en el disco.")

    html_raw = req.content_html if req.content_html is not None else (req.html if req.html is not None else "")

    try:
        sync_html_to_docx(html_raw, docx_path)

        # Invalidar caché PDF de esta plantilla
        resolved_key = str(Path(docx_path).resolve())
        if resolved_key in PDF_PREVIEW_CACHE:
            old_c = PDF_PREVIEW_CACHE.pop(resolved_key)
            old_file = TEMP_DIR / old_c.get("filename", "")
            if old_file.exists():
                try:
                    os.remove(old_file)
                except Exception:
                    pass

        return {
            "success": True,
            "filepath": docx_path,
            "message": "Archivo Word guardado exitosamente en disco con todos sus formatos preservados."
        }
    except PermissionError:
        raise HTTPException(
            status_code=409,
            detail="No se pudo guardar el archivo Word porque está abierto en Microsoft Word u otro programa. Por favor ciérralo e intenta nuevamente."
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error al guardar plantilla Word: {str(e)}")


@app.post("/api/add-docx-template")
async def add_docx_template(project_id: str = Form(...), file: UploadFile = File(...)):
    """Agrega una nueva plantilla Word a un proyecto existente."""
    p_data = load_projects_data()
    selected = None
    for p in p_data.get("projects", []):
        if p["id"] == project_id:
            selected = p
            break
    if not selected:
        raise HTTPException(status_code=404, detail="Proyecto no encontrado.")

    filename_lower = file.filename.lower()
    if not (filename_lower.endswith(".docx") or filename_lower.endswith(".doc")):
        raise HTTPException(status_code=400, detail="Formato no soportado. Debe ser .docx o .doc")

    proj_dir = Path(selected["folder_path"])
    orig_name = re.sub(r'[<>:"/\\|?*]', '_', file.filename.strip())
    saved_path = proj_dir / orig_name

    with open(saved_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    valid_docx = ensure_valid_docx(saved_path)
    html_content = docx_to_editor_html(str(valid_docx.resolve()))

    paths = selected.get("docx_paths", [])
    str_path = str(valid_docx.resolve())
    if str_path not in paths:
        paths.append(str_path)
    selected["docx_paths"] = paths
    selected["docx_path"] = str_path
    save_projects_data(p_data)

    return {
        "success": True,
        "filename": valid_docx.name,
        "filepath": str_path,
        "html": html_content,
        "docx_paths": paths
    }


@app.post("/api/delete-docx-template")
async def delete_docx_template(req: DeleteTemplateRequest):
    """Elimina físicamente una plantilla Word de un proyecto y elimina sus previsualizaciones PDF asociadas."""
    p_data = load_projects_data()
    selected = None
    for p in p_data.get("projects", []):
        if p["id"] == req.project_id:
            selected = p
            break
    if not selected:
        raise HTTPException(status_code=404, detail="Proyecto no encontrado.")

    # 1. Liberar locks COM en Word si el archivo está abierto
    folder_path = selected.get("folder_path")
    if folder_path:
        force_close_project_files(folder_path)

    # 2. Limpiar cache y archivo físico de previsualización PDF
    resolved_key = str(Path(req.docx_path).resolve()) if req.docx_path else ""
    if resolved_key and resolved_key in PDF_PREVIEW_CACHE:
        cached_info = PDF_PREVIEW_CACHE.pop(resolved_key)
        old_preview = TEMP_DIR / cached_info.get("filename", "")
        if old_preview.exists():
            try:
                os.remove(old_preview)
            except Exception:
                pass

    # 3. Eliminar físicamente el archivo .docx del disco
    if req.docx_path and os.path.exists(req.docx_path):
        try:
            os.remove(req.docx_path)
        except Exception as e:
            print(f"Aviso al eliminar plantilla Word {req.docx_path}: {e}")

    # 4. Actualizar listas en el proyecto
    paths = selected.get("docx_paths", [])
    if req.docx_path in paths:
        paths.remove(req.docx_path)
    selected["docx_paths"] = paths
    if selected.get("docx_path") == req.docx_path:
        selected["docx_path"] = paths[0] if paths else None

    # 5. Si ya no quedan plantillas, limpiar caché y cualquier archivo preview_*.pdf remanente en temp_uploads
    if not paths:
        PDF_PREVIEW_CACHE.clear()
        for f in TEMP_DIR.glob("preview_*.pdf"):
            try:
                os.remove(f)
            except Exception:
                pass

    save_projects_data(p_data)
    return {"success": True, "docx_paths": paths, "remaining": len(paths)}


def process_generation_job(job_id: str, excel_path: str, docx_paths: List[str], filename_pattern: str, output_dir_override: Optional[str] = None, settings: Optional[dict] = None, action: str = "both", selected_row_indices: Optional[List[int]] = None, generate_mode: str = "all"):
    """Proceso sincrónico en segundo plano para generar los documentos, PDFs e integrar con Outlook independientemente."""
    try:
        jobs_status[job_id] = {
            "status": "processing",
            "progress": 0,
            "total": 0,
            "current": 0,
            "output_dir": "",
            "error": None,
            "warnings": [],
            "has_warnings": False,
            "canceled": False
        }

        if excel_path.lower().endswith((".xlsx", ".xls", ".xlsm")):
            full_df = pd.read_excel(excel_path, engine='openpyxl' if excel_path.lower().endswith(('.xlsx', '.xlsm')) else None)
        else:
            full_df = pd.read_csv(excel_path)

        st = settings or {}
        output_format = st.get("output_format", "docx")
        outlook_enabled = st.get("outlook_enabled", False)
        email_col = st.get("email_column", "Email")
        check_col = st.get("check_column", "Enviado")
        send_mode = st.get("send_mode", "display")
        attachment_type = st.get("attachment_type", "pdf")
        email_subj_template = st.get("email_subject", "Notificación Oficial")
        email_body_source = st.get("email_body_source", "text")
        email_body_template = st.get("email_body", "Estimado/a {{ Nombre }},\n\nAdjuntamos su documento.")
        email_body_template_path = st.get("email_body_template", "").strip()
        single_email_override = st.get("single_email_override", "").strip()
        custom_rules = st.get("custom_paragraph_rules", [])

        # Nuevas configuraciones
        doc_check_col = st.get("doc_check_column", "Doc_Generado").strip()
        doc_check_mode = st.get("doc_check_value_mode", "timestamp")
        doc_check_custom = st.get("doc_check_custom_text", "Generado")
        all_fields_mandatory = st.get("all_fields_mandatory", False)
        mandatory_columns = st.get("mandatory_columns", [])
        template_rules_enabled = st.get("template_rules_enabled", False)
        template_rules = st.get("template_rules", [])

        # Garantizar que las columnas de control existan
        if check_col not in full_df.columns:
            full_df[check_col] = ""
        full_df[check_col] = full_df[check_col].astype(object)

        if doc_check_col and doc_check_col not in full_df.columns:
            full_df[doc_check_col] = ""
        if doc_check_col:
            full_df[doc_check_col] = full_df[doc_check_col].astype(object)

        if selected_row_indices and isinstance(selected_row_indices, list) and len(selected_row_indices) > 0:
            target_indices = [i for i in selected_row_indices if 0 <= i < len(full_df)]
        else:
            target_indices = list(range(len(full_df)))

        # Filtrar si generate_mode == "missing_only" (Solo generar faltantes)
        if generate_mode == "missing_only" and doc_check_col in full_df.columns:
            filtered_target = []
            for i in target_indices:
                val = str(full_df.iloc[i].get(doc_check_col, "")).strip()
                if not val or val.lower() in ["none", "nan", "false", "0", "no"]:
                    filtered_target.append(i)
            target_indices = filtered_target

        total_records = len(target_indices)
        jobs_status[job_id]["total"] = total_records

        if total_records == 0:
            jobs_status[job_id]["status"] = "completed"
            jobs_status[job_id]["warnings"] = ["No se encontraron documentos pendientes por generar (todas las filas están al día)."]
            jobs_status[job_id]["has_warnings"] = True
            return

        if output_dir_override and os.path.exists(output_dir_override):
            job_output_dir = Path(output_dir_override)
        else:
            timestamp_folder = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
            job_output_dir = OUTPUT_BASE_DIR / f"generacion_{timestamp_folder}"
            job_output_dir.mkdir(exist_ok=True)

        jobs_status[job_id]["output_dir"] = str(job_output_dir.resolve())

        need_pdf = (action in ["generate_docs", "both"] and output_format in ["pdf", "both"]) or (action in ["send_emails", "both"] and outlook_enabled and attachment_type in ["pdf", "both"])
        
        import pythoncom
        pythoncom.CoInitialize()
        job_warnings = []

        for step_idx, orig_idx in enumerate(target_indices):
            if jobs_status.get(job_id, {}).get("canceled"):
                jobs_status[job_id]["status"] = "canceled"
                pythoncom.CoUninitialize()
                return

            row = full_df.iloc[orig_idx]
            context = row.to_dict()
            context["_index"] = step_idx + 1
            context["_row"] = orig_idx + 1

            clean_context = {}
            for k, v in context.items():
                k_str = str(k).strip()
                if pd.isna(v):
                    clean_context[k_str] = ""
                else:
                    clean_context[k_str] = str(v)

            # Validar Campos Obligatorios antes de generar la fila
            if all_fields_mandatory:
                missing = [c for c in full_df.columns if c not in ["_index", "_row", check_col, doc_check_col] and not str(clean_context.get(c, "")).strip()]
                if missing:
                    job_warnings.append(f"Fila {orig_idx + 1}: Omitida por tener campos obligatorios vacíos ({', '.join(missing)})")
                    continue
            elif mandatory_columns and isinstance(mandatory_columns, list) and len(mandatory_columns) > 0:
                missing = [c for c in mandatory_columns if not str(clean_context.get(c, "")).strip()]
                if missing:
                    job_warnings.append(f"Fila {orig_idx + 1}: Omitida por falta de campos obligatorios ({', '.join(missing)})")
                    continue

            if custom_rules:
                for rule in custom_rules:
                    rule_col = str(rule.get("column", "")).strip()
                    rule_val = str(rule.get("value", "")).strip()
                    custom_text = rule.get("custom_text", "")

                    if rule_col in clean_context and str(clean_context[rule_col]).strip().lower() == rule_val.lower():
                        tmpl = Template(custom_text)
                        rendered_text = tmpl.render(clean_context)
                        clean_context[rule_col] = rendered_text
                        clean_context[f"p_{rule_col}"] = rendered_text
                        clean_context[f"p_{rule_col.lower()}"] = rendered_text

            # Evaluar Reglas de Plantilla Condicionales ("Reglas de Plantilla")
            active_docx_paths = list(docx_paths)
            if template_rules_enabled and template_rules and isinstance(template_rules, list):
                matched_paths = []
                for rule in template_rules:
                    r_col = str(rule.get("column", "")).strip()
                    r_val = str(rule.get("value", "")).strip()
                    r_tpl = str(rule.get("template_filename", "")).strip()
                    if r_col in clean_context and str(clean_context[r_col]).strip().lower() == r_val.lower():
                        for dp in docx_paths:
                            if Path(dp).name.lower() == Path(r_tpl).name.lower() or str(Path(dp).resolve()).lower() == str(Path(r_tpl).resolve()).lower():
                                if dp not in matched_paths:
                                    matched_paths.append(dp)
                if matched_paths:
                    active_docx_paths = matched_paths

            generated_docx_files = []
            generated_pdf_files = []

            # Loop a través de las plantillas activas para esta fila
            for tpl_idx, d_path in enumerate(active_docx_paths):
                if not os.path.exists(d_path):
                    continue

                try:
                    doc = DocxTemplate(d_path)
                    doc.render(clean_context)
                except Exception as render_ex:
                    jobs_status[job_id]["status"] = "error"
                    jobs_status[job_id]["error"] = f"Error al procesar plantilla '{Path(d_path).name}' en la fila {orig_idx + 1}: {str(render_ex)}"
                    pythoncom.CoUninitialize()
                    return

                file_name = filename_pattern
                for k, v in clean_context.items():
                    file_name = file_name.replace(f"{{{{ {k} }}}}", str(v))
                    file_name = file_name.replace(f"{{{{{k}}}}}", str(v))
                    file_name = file_name.replace(f"{{ {k} }}", str(v))
                    file_name = file_name.replace(f"{{{k}}}", str(v))

                file_name = file_name.replace("{{index}}", str(step_idx + 1)).replace("{{ index }}", str(step_idx + 1))
                file_name = file_name.replace("{index}", str(step_idx + 1))

                if len(active_docx_paths) > 1:
                    tpl_stem = Path(d_path).stem
                    base_name, ext = os.path.splitext(file_name)
                    if not ext:
                        ext = ".docx"
                    file_name = f"{base_name}_{tpl_stem}{ext}"

                invalid_chars = '<>:"/\\|?*'
                for char in invalid_chars:
                    file_name = file_name.replace(char, "_")

                if not file_name.endswith(".docx"):
                    file_name += ".docx"

                docx_output_path = job_output_dir / file_name
                
                if action in ["generate_docs", "both", "send_emails"]:
                    doc.save(str(docx_output_path))
                    generated_docx_files.append(docx_output_path)

                pdf_output_path = docx_output_path.with_suffix(".pdf")

                if need_pdf:
                    try:
                        word_app = win32com.client.DispatchEx("Word.Application")
                        word_app.Visible = False
                        word_app.DisplayAlerts = 0
                        wdoc = word_app.Documents.Open(str(docx_output_path.resolve()), ReadOnly=True, ConfirmConversions=False)
                        wdoc.SaveAs2(str(pdf_output_path.resolve()), FileFormat=17) # 17 = wdFormatPDF
                        wdoc.Close(0)
                        word_app.Quit()
                        generated_pdf_files.append(pdf_output_path)
                    except Exception as ex_pdf:
                        print("Error generando PDF con Word COM:", ex_pdf)
                        try:
                            from docx2pdf import convert
                            convert(str(docx_output_path.resolve()), str(pdf_output_path.resolve()))
                            if pdf_output_path.exists():
                                generated_pdf_files.append(pdf_output_path)
                        except Exception as ex_fallback:
                            job_warnings.append(f"Fila {orig_idx + 1}: No se pudo convertir a PDF ({str(ex_pdf)})")

                if action == "generate_docs" and output_format == "pdf" and docx_output_path.exists() and pdf_output_path.exists():
                    try:
                        os.remove(docx_output_path)
                    except Exception:
                        pass

            # REGISTRAR COLUMNA DE CONTROL DE GENERACIÓN DE DOCUMENTOS
            if action in ["generate_docs", "both"] and (generated_docx_files or generated_pdf_files) and doc_check_col:
                if doc_check_mode == "timestamp":
                    doc_stamp = pd.Timestamp.now().strftime("%d/%m/%Y %H:%M:%S")
                elif doc_check_mode == "format":
                    doc_stamp = "Word (.docx)" if output_format == "docx" else ("PDF (.pdf)" if output_format == "pdf" else "Word + PDF")
                else:
                    doc_stamp = doc_check_custom if doc_check_custom else "Generado"

                full_df.at[orig_idx, doc_check_col] = doc_stamp
                if excel_path.lower().endswith((".xlsx", ".xls", ".xlsm")):
                    try:
                        is_xlsm = excel_path.lower().endswith(".xlsm")
                        wb_gen = openpyxl.load_workbook(excel_path, keep_vba=is_xlsm)
                        ws_gen = wb_gen.active
                        h_cols = [ws_gen.cell(row=1, column=c).value for c in range(1, ws_gen.max_column + 1)]
                        if doc_check_col in h_cols:
                            c_num = h_cols.index(doc_check_col) + 1
                            ws_gen.cell(row=orig_idx + 2, column=c_num).value = doc_stamp
                            wb_gen.save(excel_path)
                        else:
                            full_df.to_excel(excel_path, index=False)
                    except Exception:
                        full_df.to_excel(excel_path, index=False)

            # 3. ENVÍO DE CORREO POR OUTLOOK
            if action in ["send_emails", "both"]:
                target_email = clean_context.get(email_col, "").strip()
                check_val = str(clean_context.get(check_col, "")).strip()
                is_already_sent = bool(check_val) and check_val.upper() not in ["NO", "FALSE", "0", "NONE"]

                # Obtener destinatarios finales (principal + campos adicionales de la tabla + adicionales condicionales o fijos)
                recipients = []
                if single_email_override:
                    recipients.append(single_email_override)
                else:
                    # 1. Columna de correo principal
                    if target_email:
                        for em in re.split(r'[;,]', target_email):
                            em_clean = em.strip()
                            if em_clean and em_clean not in recipients:
                                recipients.append(em_clean)
                    
                    # 2. Campos adicionales de correo configurados desde columnas de la tabla (ej: Correo Estudiantes, Correo Profesores)
                    extra_email_cols = st.get("email_columns", [])
                    for ecol in extra_email_cols:
                        col_key = ecol.get("column", "").strip() if isinstance(ecol, dict) else str(ecol).strip()
                        if col_key:
                            col_val = str(clean_context.get(col_key, "")).strip()
                            if col_val:
                                for em in re.split(r'[;,]', col_val):
                                    em_clean = em.strip()
                                    if em_clean and em_clean not in recipients:
                                        recipients.append(em_clean)

                    # 3. Correos adicionales/fijos con reglas
                    additional_emails_config = st.get("additional_emails", [])
                    for add_item in additional_emails_config:
                        add_email = add_item.get("email", "").strip()
                        if not add_email:
                            continue
                        
                        rule_col = add_item.get("rule_column", "").strip()
                        rule_val = add_item.get("rule_value", "").strip()
                        
                        if not rule_col:
                            # Permanente
                            if add_email not in recipients:
                                recipients.append(add_email)
                        else:
                            # Condicional
                            row_val = str(clean_context.get(rule_col, "")).strip()
                            if row_val.lower() == rule_val.lower():
                                if add_email not in recipients:
                                    recipients.append(add_email)

                final_recipients = "; ".join(recipients)

                if final_recipients and (not is_already_sent or single_email_override or selected_row_indices is not None):
                    try:
                        outlook = win32com.client.DispatchEx("Outlook.Application")
                        mail = outlook.CreateItem(0) # 0 = olMailItem
                        mail.To = final_recipients

                        subj_tmpl = Template(email_subj_template)
                        mail.Subject = subj_tmpl.render(clean_context)

                        if email_body_source in ["template", "word", "docx"]:
                            target_docx_for_body = email_body_template_path
                            if not target_docx_for_body or not os.path.exists(target_docx_for_body):
                                if generated_docx_files and os.path.exists(generated_docx_files[0]):
                                    target_docx_for_body = str(generated_docx_files[0])
                                elif docx_paths and os.path.exists(docx_paths[0]):
                                    target_docx_for_body = docx_paths[0]

                            if target_docx_for_body and os.path.exists(target_docx_for_body):
                                try:
                                    doc_body = DocxTemplate(target_docx_for_body)
                                    doc_body.render(clean_context)
                                    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp_b:
                                        doc_body.save(tmp_b.name)
                                        tmp_b_path = tmp_b.name

                                    with open(tmp_b_path, "rb") as w_file:
                                        html_conv = mammoth.convert_to_html(w_file)
                                        mail.HTMLBody = html_conv.value
                                    try:
                                        os.remove(tmp_b_path)
                                    except Exception:
                                        pass
                                except Exception as ex_html:
                                    body_tmpl = Template(email_body_template)
                                    mail.Body = body_tmpl.render(clean_context)
                            else:
                                body_tmpl = Template(email_body_template)
                                mail.Body = body_tmpl.render(clean_context)
                        else:
                            body_tmpl = Template(email_body_template)
                            mail.Body = body_tmpl.render(clean_context)

                        if attachment_type == "pdf":
                            if generated_pdf_files:
                                single_pdf = generated_pdf_files[-1]
                                if single_pdf.exists():
                                    mail.Attachments.Add(str(single_pdf.resolve()))
                        elif attachment_type in ["docx", "word"]:
                            if generated_docx_files:
                                single_docx = generated_docx_files[-1]
                                if single_docx.exists():
                                    mail.Attachments.Add(str(single_docx.resolve()))
                        elif attachment_type == "both":
                            attached_paths = set()
                            for pf in generated_pdf_files:
                                if pf.exists() and str(pf) not in attached_paths:
                                    mail.Attachments.Add(str(pf.resolve()))
                                    attached_paths.add(str(pf))
                            for df_f in generated_docx_files:
                                if df_f.exists() and str(df_f) not in attached_paths:
                                    mail.Attachments.Add(str(df_f.resolve()))
                                    attached_paths.add(str(df_f))

                        if send_mode == "send":
                            mail.Send()
                        else:
                            mail.Display()

                        if not single_email_override:
                            now_str = pd.Timestamp.now().strftime("%d/%m/%Y %H:%M:%S")
                            full_df.at[orig_idx, check_col] = now_str
                            if excel_path.lower().endswith((".xlsx", ".xls", ".xlsm")):
                                try:
                                    is_xlsm = excel_path.lower().endswith(".xlsm")
                                    wb_gen = openpyxl.load_workbook(excel_path, keep_vba=is_xlsm)
                                    ws_gen = wb_gen.active
                                    header_cols = [ws_gen.cell(row=1, column=c).value for c in range(1, ws_gen.max_column + 1)]
                                    if check_col in header_cols:
                                        c_num = header_cols.index(check_col) + 1
                                        ws_gen.cell(row=orig_idx + 2, column=c_num).value = now_str
                                        wb_gen.save(excel_path)
                                    else:
                                        full_df.to_excel(excel_path, index=False)
                                except Exception:
                                    full_df.to_excel(excel_path, index=False)
                            else:
                                full_df.to_csv(excel_path, index=False)

                    except Exception as ex_mail:
                        print("Error enviando correo Outlook:", ex_mail)
                        job_warnings.append(f"Fila {orig_idx + 1} ({final_recipients}): {str(ex_mail)}")

            progress_pct = int(((step_idx + 1) / total_records) * 100)
            jobs_status[job_id]["current"] = step_idx + 1
            jobs_status[job_id]["progress"] = progress_pct

        pythoncom.CoUninitialize()
        jobs_status[job_id]["status"] = "completed"
        if job_warnings:
            jobs_status[job_id]["has_warnings"] = True
            jobs_status[job_id]["warnings"] = job_warnings

    except Exception as e:
        jobs_status[job_id]["status"] = "error"
        jobs_status[job_id]["error"] = str(e)


@app.post("/api/generate")
async def start_generate(req: GenerateRequest, background_tasks: BackgroundTasks):
    """Inicia la generación masiva de documentos y/o envío independiente de correos."""
    if not os.path.exists(req.excel_path):
        raise HTTPException(status_code=400, detail="El archivo Excel especificado no existe.")
    
    docx_paths = req.docx_paths or ([req.docx_path] if req.docx_path else [])
    if not docx_paths or not any(os.path.exists(p) for p in docx_paths):
        raise HTTPException(status_code=400, detail="No se especificó ninguna plantilla Word válida.")

    job_id = f"job_{uuid.uuid4().hex[:10]}"
    pattern = req.output_filename_pattern or "doc_{{index}}.docx"

    background_tasks.add_task(
        process_generation_job,
        job_id=job_id,
        excel_path=req.excel_path,
        docx_paths=docx_paths,
        filename_pattern=pattern,
        output_dir_override=req.output_dir_override,
        settings=req.settings,
        action=req.action or "both",
        selected_row_indices=req.selected_row_indices,
        generate_mode=req.generate_mode or "all"
    )

    return {"success": True, "job_id": job_id}


@app.post("/api/cancel-job")
async def cancel_job(req: CancelJobRequest):
    """Cancela un trabajo de generación en segundo plano inmediatamente."""
    if req.job_id in jobs_status:
        jobs_status[req.job_id]["canceled"] = True
        jobs_status[req.job_id]["status"] = "canceled"
        return {"success": True, "message": "Proceso cancelado."}
    return {"success": False, "message": "Tarea no encontrada."}


@app.get("/api/progress/{job_id}")
async def get_progress(job_id: str):
    """Retorna el estado de avance de una tarea de generación."""
    if job_id not in jobs_status:
        return {"status": "not_found", "progress": 0}
    return jobs_status[job_id]


@app.post("/api/open-file")
async def open_file(data: dict):
    """Abre un archivo nativo (Excel, Word) en su aplicación predeterminada del SO."""
    filepath = data.get("filepath")
    if not filepath or not os.path.exists(filepath):
        raise HTTPException(status_code=404, detail="El archivo especificado no existe en el disco.")

    try:
        abs_path = os.path.normpath(os.path.abspath(filepath))
        if sys.platform == "win32":
            try:
                os.startfile(abs_path)
            except Exception:
                subprocess.Popen(f'explorer "{abs_path}"', shell=True)
        elif sys.platform == "darwin":
            subprocess.run(["open", abs_path])
        else:
            subprocess.run(["xdg-open", abs_path])
        return {"success": True}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error al abrir archivo: {str(e)}")


@app.post("/api/open-output-folder")
async def open_output_folder(path: dict):
    """Abre la carpeta de salida en el Explorador de archivos nativo de Windows."""
    folder_path = path.get("folder_path")
    if not folder_path or not os.path.exists(folder_path):
        raise HTTPException(status_code=404, detail="La carpeta del proyecto no existe en el disco.")

    try:
        abs_folder = os.path.normpath(os.path.abspath(folder_path))
        if sys.platform == "win32":
            try:
                os.startfile(abs_folder)
            except Exception:
                subprocess.Popen(f'explorer "{abs_folder}"', shell=True)
        elif sys.platform == "darwin":
            subprocess.run(["open", abs_folder])
        else:
            subprocess.run(["xdg-open", abs_folder])
        return {"success": True}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error al abrir carpeta: {str(e)}")


@app.post("/api/create-sample-files")
async def create_sample_files():
    """Genera automáticamente un Excel y un DOCX de muestra para probar la app al instante."""
    try:
        sample_dir = BASE_DIR / "sample_data"
        sample_dir.mkdir(exist_ok=True)

        excel_path = sample_dir / "listado_clientes.xlsx"
        docx_path = sample_dir / "comunicado_oficial.docx"

        df_sample = pd.DataFrame([
            {"Nombre": "María García", "Empresa": "Soluciones Alpha S.A.", "Cargo": "Gerente General", "Concepto": "Servicios Profesionales", "Monto": "$1,500.00", "Fecha": "11/08/2026", "Email": "mgarcia@ejemplo.com", "Enviado": "NO"},
            {"Nombre": "Carlos Mendoza", "Empresa": "Consultora Beta", "Cargo": "Director de Operaciones", "Concepto": "Consultoría Mensual", "Monto": "$2,800.00", "Fecha": "11/08/2026", "Email": "cmendoza@ejemplo.com", "Enviado": "NO"},
            {"Nombre": "Ana Belén Torres", "Empresa": "Innovación Gamma", "Cargo": "Jefa de Proyectos", "Concepto": "Mantenimiento Técnico", "Monto": "$950.00", "Fecha": "11/08/2026", "Email": "atorres@ejemplo.com", "Enviado": "NO"},
        ])
        df_sample.to_excel(excel_path, index=False)

        doc = Document()
        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_title = p_title.add_run("COMUNICADO OFICIAL DE NOTIFICACIÓN")
        run_title.font.name = "Arial"
        run_title.font.size = Pt(20)
        run_title.font.bold = True
        run_title.font.color.rgb = RGBColor(30, 58, 138)

        doc.add_paragraph()

        p_body = doc.add_paragraph()
        p_body.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run_body1 = p_body.add_run("Por medio del presente documento, notificamos formalmente a ")
        run_body1.font.size = Pt(13)

        run_var_nombre = p_body.add_run("{{ Nombre }}")
        run_var_nombre.font.size = Pt(14)
        run_var_nombre.font.bold = True
        run_var_nombre.font.color.rgb = RGBColor(194, 65, 12)

        run_body2 = p_body.add_run(", quien se desempeña como ")
        run_body2.font.size = Pt(13)

        run_var_cargo = p_body.add_run("{{ Cargo }}")
        run_var_cargo.font.size = Pt(14)
        run_var_cargo.font.bold = True

        run_body3 = p_body.add_run(" en la empresa ")
        run_body3.font.size = Pt(13)

        run_var_empresa = p_body.add_run("{{ Empresa }}")
        run_var_empresa.font.size = Pt(13)
        run_var_empresa.font.bold = True

        run_body4 = p_body.add_run(", respecto al concepto de ")
        run_body4.font.size = Pt(13)

        run_var_concepto = p_body.add_run("{{ Concepto }}")
        run_var_concepto.font.size = Pt(13)
        run_var_concepto.font.bold = True

        run_body5 = p_body.add_run(" por el valor de ")
        run_body5.font.size = Pt(13)

        run_var_monto = p_body.add_run("{{ Monto }}")
        run_var_monto.font.size = Pt(13)
        run_var_monto.font.bold = True

        run_body6 = p_body.add_run(" emitido en fecha ")
        run_body6.font.size = Pt(13)

        run_var_fecha = p_body.add_run("{{ Fecha }}")
        run_var_fecha.font.size = Pt(13)
        run_var_fecha.font.bold = True

        doc.add_paragraph()

        p_firma = doc.add_paragraph()
        p_firma.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run_f1 = p_firma.add_run("Emitido por: Dirección General de Operaciones")
        run_f1.font.size = Pt(12)

        doc.save(str(docx_path))

        return {
            "success": True,
            "excel_url": "/api/download-sample/listado_clientes.xlsx",
            "docx_url": "/api/download-sample/comunicado_oficial.docx",
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error creando archivos muestra: {str(e)}")


@app.get("/api/download-sample/{filename}")
async def download_sample(filename: str):
    sample_file = BASE_DIR / "sample_data" / filename
    if sample_file.exists():
        return FileResponse(path=str(sample_file), filename=filename)
    raise HTTPException(status_code=404, detail="Archivo no encontrado")


@app.post("/api/reload-docx")
async def reload_docx(data: dict):
    """Recarga el contenido de un archivo DOCX desde disco y lo devuelve convertido a HTML."""
    filepath = data.get("docx_path")
    if not filepath or not os.path.exists(filepath):
        raise HTTPException(status_code=404, detail="El archivo Word especificado no existe en el disco.")

    try:
        html_content = docx_to_editor_html(filepath)

        return {
            "success": True,
            "filename": os.path.basename(filepath),
            "html": html_content
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error leyendo el archivo Word: {str(e)}")
