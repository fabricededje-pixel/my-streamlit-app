import base64
import json
import os
import zipfile

from pypdf import PdfReader, PdfWriter

from exporters.json_store import profile_from_data, profile_to_dict


DOCX_PAYLOAD_PATH = "cv_builder/profile.json"
PDF_PAYLOAD_KEY = "/CVBuilderProfile"


def _serialize_profile(profile) -> str:
    raw = json.dumps(profile_to_dict(profile), ensure_ascii=False).encode("utf-8")
    return base64.b64encode(raw).decode("ascii")


def _deserialize_profile(payload: str):
    raw = base64.b64decode(payload.encode("ascii")).decode("utf-8")
    return profile_from_data(json.loads(raw))


def embed_profile_in_docx(docx_path: str, profile) -> None:
    payload = json.dumps(profile_to_dict(profile), ensure_ascii=False, indent=2)
    with zipfile.ZipFile(docx_path, "a", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(DOCX_PAYLOAD_PATH, payload)


def extract_profile_from_docx(docx_path: str):
    with zipfile.ZipFile(docx_path, "r") as archive:
        try:
            payload = archive.read(DOCX_PAYLOAD_PATH).decode("utf-8")
        except KeyError as exc:
            raise ValueError("Keine eingebetteten CV-Daten in dieser DOCX-Datei gefunden.") from exc
    return profile_from_data(json.loads(payload))


def embed_profile_in_pdf(pdf_path: str, profile) -> None:
    reader = PdfReader(pdf_path)
    writer = PdfWriter()

    for page in reader.pages:
        writer.add_page(page)

    metadata = {}
    if reader.metadata:
        for key, value in reader.metadata.items():
            if value is not None:
                metadata[key] = str(value)

    metadata[PDF_PAYLOAD_KEY] = _serialize_profile(profile)
    writer.add_metadata(metadata)

    temp_path = pdf_path + ".tmp"
    with open(temp_path, "wb") as file_obj:
        writer.write(file_obj)

    os.replace(temp_path, pdf_path)


def extract_profile_from_pdf(pdf_path: str):
    reader = PdfReader(pdf_path)
    metadata = reader.metadata or {}
    payload = metadata.get(PDF_PAYLOAD_KEY)
    if not payload:
        raise ValueError("Keine eingebetteten CV-Daten in dieser PDF-Datei gefunden.")
    return _deserialize_profile(payload)
