"""Shared observation/specimen VWorld workflow; no credential diagnostics."""

from PySide6.QtWidgets import QDialog, QMessageBox

from services.api_key_service import ApiKeyError, ApiKeyService
from services.vworld_geocoding_service import VWorldGeocodingService
from services.address_mapping_service import apply_address
from ui.dialogs.api_settings_dialog import ApiSettingsDialog
from ui.dialogs.address_lookup_dialog import AddressLookupDialog, AddressPreviewDialog


def lookup_addresses(parent, targets, specimen=False, differences_only=False):
    if not targets:
        return {}
    keys = ApiKeyService()
    try:
        key, expiry = keys.load()
        if not key:
            ApiSettingsDialog(parent, key_service=keys).exec()
            key, expiry = keys.load()
        if not key:
            return {}
    except ApiKeyError as exc:
        QMessageBox.warning(parent, "VWorld API 키", str(exc))
        return {}
    dialog = AddressLookupDialog(VWorldGeocodingService(), key, targets, parent)
    key = ""
    if dialog.exec() != QDialog.Accepted:
        return {}
    preview = AddressPreviewDialog(targets, dialog.worker.results, dialog.worker.errors, specimen, parent,
                                   dialog.worker.references, dialog.worker.reference_errors, differences_only)
    if preview.exec() != QDialog.Accepted:
        return {}
    return preview.planned_records()
