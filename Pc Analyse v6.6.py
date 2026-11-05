"""
PC-Analyse-Tool v6.5 (Automatische Spracherkennung DE/EN)
Zeigt auf einen Blick, wo ein Windows-PC ausgebremst wird:
Speicherplatz      - Freier Speicher, Laufwerkstyp und verfügbare Windows-Laufwerksdaten
Autostart          - Programme, die beim Anmelden starten, inkl. Prüfung auf fehlende Dateien
Größte Ordner      - Ordner, Dateien und Dateitypen nach Platzverbrauch, mit Datum und Alter

Das Tool ändert nichts am System. Es liest nur und zeigt an.
Läuft ohne zusätzliche Pakete (nur Python-Standardbibliothek).
"""
import csv
import heapq
import json
import locale
import os
import queue
import re
import shutil
import stat
import subprocess
import threading
import time
import traceback
import tkinter as tk
import tkinter.font as tkfont
from collections import defaultdict
from concurrent.futures import FIRST_COMPLETED, ThreadPoolExecutor, wait
from tkinter import filedialog, messagebox, ttk

try:
    import winreg
except ImportError:
    winreg = None

# ----------------------------------------------------------------------------
# Automatische Spracherkennung
# ----------------------------------------------------------------------------
def detect_language():
    """Erkennt die Systemsprache. Fallback ist Deutsch."""
    try:
        sys_lang = locale.getdefaultlocale()[0]
        if sys_lang and sys_lang.lower().startswith("en"):
            return "en"
    except Exception:
        pass
    return "de"

LANG = detect_language()

TEXTS = {
    "de": {
        "app_title": "PC-Analyse-Tool V6.6",
        "header_title": "PC-Analyse",
        "header_subtitle": "Wo geht der Speicherplatz hin und was bremst den PC? Das Tool zeigt nur an und verändert nichts.",
        "btn_safety": "Sicherheitshinweise",
        "btn_cleanup": "Sicher aufräumen",
        "tab_storage": "Speicherplatz",
        "tab_autostart": "Autostart",
        "tab_folders": "Größte Ordner & Dateien",
        "btn_refresh": "Aktualisieren",
        "btn_drive_opt": "Windows-Laufwerksoptimierung öffnen",
        "lbl_trim_checking": "TRIM-Einstellung wird geprüft …",
        "lbl_trim_result": "TRIM-Einstellung von Windows: {}. Das ist eine Windows-Einstellung und kein direkter Funktionstest des Laufwerks.",
        "lbl_storage_warn": "Warnung, wenn weniger als {} % frei sind. Mit 'Analysieren' siehst du, was den Platz belegt. Die Laufwerksoptimierung öffnet nur das Windows-Werkzeug.",
        "lbl_no_drives": "Keine Laufwerke gefunden.",
        "lbl_drives_summary": "{} {} · {} gesamt · {} belegt · {} frei",
        "lbl_drives_low": " · ⚠ {} mit wenig Platz",
        "unit_drive": "Laufwerk",
        "unit_drives": "Laufwerke",
        "chip_ok": "In Ordnung",
        "chip_warn": "Wenig Platz",
        "chip_crit": "Fast voll",
        "lbl_free_of": "{} frei von {}",
        "lbl_used_free": "{} belegt · {} % frei",
        "lbl_type": "Typ:",
        "lbl_model": "Modell:",
        "lbl_interface": "Anschluss:",
        "lbl_health": "Gemeldeter Zustand:",
        "lbl_temp_wear": "Temperatur/Verschleiß: von Windows nicht bereitgestellt",
        "lbl_temperature": "Temperatur",
        "lbl_wear": "Verschleiß",
        "health_good": "Gut",
        "health_warning": "Warnung",
        "health_critical": "Kritisch",
        "health_unknown": "Unbekannt",
        "btn_analyze": "Speicherplatz analysieren",
        "btn_choose": "Ordner wählen & analysieren…",
        "btn_up": "↑ Eine Ebene höher",
        "btn_cancel": "Abbrechen",
        "btn_export": "Exportieren…",
        "lbl_folder_not_chosen": "Noch kein Ordner gewählt",
        "lbl_folder_hint": "Wähle einen Ordner oder ein Laufwerk, um die Analyse zu starten. Klick auf eine Spaltenüberschrift sortiert, Rechtsklick öffnet ein Menü.",
        "lbl_no_analysis": "Noch keine Analyse.\n\nWähle oben einen Ordner oder starte im Reiter 'Speicherplatz' die Analyse eines Laufwerks.",
        "lbl_analyzing": "Analysiere… Das kann bei großen Ordnern etwas dauern.",
        "lbl_canceling": "Breche ab...",
        "lbl_analysis_aborted": "Analyse abgebrochen.",
        "lbl_analysis_failed": "Die Analyse ist fehlgeschlagen.",
        "lbl_finished": "Fertig. {} in {} Dateien.",
        "lbl_not_readable": " {} Einträge nicht lesbar (fehlende Rechte), die Summe kann daher etwas zu klein sein.",
        "lbl_active_entries": "{} aktive Einträge von {}",
        "lbl_missing_files": " · {} mit fehlender Datei (rot, oft Reste deinstallierter Software)",
        "lbl_autostart_win_only": "Die Autostart-Anzeige funktioniert nur unter Windows.",
        "lbl_autostart_only_display": "Nur Anzeige: Zum Deaktivieren nutze den Task-Manager (Reiter 'Autostart'). Deaktivieren ist umkehrbar, Löschen nicht: Entferne hier nichts aus der Registry. Die Hinweise sind Faustregeln, keine Garantie.",
        "btn_taskmgr": "Task-Manager öffnen",
        "btn_startup_settings": "Windows-Autostart-Einstellungen",
        "error_title": "Fehler",
        "error_taskmgr_failed": "Task-Manager konnte nicht gestartet werden:\n{}",
        "error_settings_failed": "Einstellungen konnten nicht geöffnet werden:\n{}",
        "error_drive_opt_failed": "Die Windows-Laufwerksoptimierung konnte nicht geöffnet werden:\n{}\n\nTipp: Starte das PC-Analyse-Tool einmalig als Administrator (Rechtsklick auf die .py-Datei oder das Icon → 'Als Administrator ausführen'), dann klappt es auch ohne UAC-Abfrage.",
        "error_not_available": "Nicht verfügbar",
        "error_drive_opt_win_only": "Die Windows-Laufwerksoptimierung gibt es nur unter Windows.",
        "safety_title": "Wichtige Sicherheitshinweise",
        "safety_warning_title": "Vorsicht: heikler Bereich",
        "safety_warning_text": "{}\n\nDu darfst dir den Bereich ansehen. Lösche, verschiebe oder ändere dort aber nichts, wenn du nicht genau weißt, was es ist. Windows oder Programme können sonst nicht mehr funktionieren.\n\nTrotzdem im Explorer öffnen?",
        "lbl_read_before_delete": "⚠  Bitte kurz lesen, bevor du etwas löschst",
        "lbl_hide_startup_notice": "Diesen Hinweis beim Start nicht mehr automatisch anzeigen",
        "btn_understood": "Verstanden",
        "error_export_failed": "Export fehlgeschlagen:\n{}",
        "error_exported": "Exportiert",
        "error_exported_text": "Ergebnisse gespeichert:\n{}",
        "error_folder_not_found": "Ordner nicht gefunden:\n{}",
        "error_unexpected": "Unerwarteter Fehler",
        "col_program": "Programm",
        "col_status": "Status",
        "col_file": "Datei",
        "col_hint": "Hinweis",
        "col_command": "Befehl / Pfad",
        "col_source": "Quelle",
        "status_active": "Aktiv",
        "status_disabled": "Deaktiviert",
        "status_present": "vorhanden",
        "status_missing": "fehlt",
        "hint_missing_file": "Programmdatei fehlt – meist Rest einer Deinstallation",
        "hint_security": "⚠ Sicherheitssoftware – besser aktiv lassen",
        "hint_windows": "⚠ Windows-Komponente – nicht ändern",
        "hint_driver": "⚠ Treiber-/Hardware-Tool – im Zweifel aktiv lassen",
        "col_folder": "Ordner",
        "col_size": "Größe",
        "col_share": "Anteil",
        "col_newest": "Neueste Änderung",
        "col_age": "Alter",
        "col_changed": "Geändert",
        "col_type": "Dateityp",
        "tab_folder": "Ordner",
        "tab_files": "Größte Dateien",
        "tab_types": "Dateitypen",
        "lbl_files_in_root": "(Dateien direkt in diesem Ordner)",
        "export_csv": "CSV (Excel)",
        "export_txt": "Textbericht",
        "export_default_name": "PC-Analyse_{}_{}.csv",
        "today": "heute",
        "days": "{} Tage",
        "months": "{} Monate",
        "years": "{} Jahre",
        "no_extension": "(ohne Endung)",
        "ssd": "SSD",
        "hdd": "Festplatte (HDD)",
        "scm": "Speicher (SCM)",
        "nvme": "NVMe",
        "sata": "SATA",
        "usb": "USB",
        "sas": "SAS",
        "ata": "ATA",
        "raid": "RAID",
        "trim_enabled": "Aktiviert",
        "trim_disabled": "Deaktiviert",
        "trim_not_available": "Nicht verfügbar",
        "open_in_explorer": "Im Explorer öffnen",
        "show_in_explorer": "Im Explorer anzeigen",
        "copy_path": "Pfad kopieren",
        "analyze_here": "Diesen Ordner analysieren (hineingehen)",
        "lbl_main_hint": "⚠  Lösche selbst nichts, was du nicht kennst, besonders nichts in Windows, in Programmordnern oder in der Registry. Heikle Bereiche sind mit ⛔ / ⚠ markiert.",
        "root_special_sysvol": "⛔ Wiederherstellungspunkte von Windows – nicht löschen",
        "root_special_recovery": "⛔ Windows-Wiederherstellung – nicht löschen",
        "root_special_winre": "⛔ Windows-Update-Wiederherstellung – nicht löschen",
        "root_special_boot": "⛔ Startdateien von Windows – nicht löschen",
        "root_special_efi": "⛔ Startdateien von Windows – nicht löschen",
        "root_special_bootmgr": "⛔ Startdatei von Windows – nicht löschen",
        "root_special_pagefile": "⛔ Auslagerungsdatei von Windows – nicht löschen",
        "root_special_swapfile": "⛔ Systemdatei von Windows – nicht löschen",
        "root_special_hiberfil": "⛔ Ruhezustand-Datei – nicht löschen",
        "root_special_recycle": "⚠ Papierkorb – über 'Papierkorb leeren' bereinigen",
        "root_special_winold": "⚠ Alter Windows-Stand – nur über die Windows-Datenträgerbereinigung entfernen",
        "root_special_winbt": "⚠ Update-Reste – nur über die Windows-Datenträgerbereinigung entfernen",
        "root_special_configmsi": "⚠ Installer-Reste – nicht manuell löschen",
        "loc_windir": "⛔ Windows-System – nicht löschen",
        "loc_progfiles": "⚠ Programmdateien – nur über 'Apps deinstallieren' entfernen",
        "loc_progdata": "⚠ Programmdaten – nicht manuell löschen",
        "loc_appdata": "⚠ App-Daten – nicht manuell löschen",
        "safety_notice": (
            "• Dieses Tool zeigt nur an und löscht nichts. Über den Explorer (Rechtsklick → Öffnen) kommst du aber "
            "direkt zu den Dateien, deshalb bitte vorsichtig sein.\n\n"
            "• Lösche niemals Dateien oder Ordner in C:\\Windows, Program Files, ProgramData oder AppData und keine "
            "Systemdateien wie pagefile.sys oder hiberfil.sys. Windows oder Programme können danach nicht mehr starten.\n\n"
            "• Große Dateien sind nicht automatisch überflüssig. Lösche nur, was du selbst kennst, zum Beispiel eigene "
            "Downloads, alte Videos oder Installationsdateien.\n\n"
            "• Programme nie durch Löschen ihres Ordners entfernen: Einstellungen → Apps (Installierte Apps bzw. "
            "Apps & Features) → Deinstallieren.\n\n"
            "• Autostart: Einträge nur deaktivieren (das ist umkehrbar), nicht löschen. Sicherheitssoftware und "
            "Treiber-Tools besser aktiv lassen.\n\n"
            "• Die Windows-Registry nie von Hand ändern und keine 'Registry-Cleaner' benutzen. Dieses Tool liest die "
            "Registry nur.\n\n"
            "• Zum sicheren Aufräumen eignet sich die Windows-Speicheroptimierung (Button 'Sicher aufräumen'). "
            "Vor größeren Aufräumaktionen ein Backup oder einen Wiederherstellungspunkt anlegen.\n\n"
            "• Im Zweifel: lieber nichts löschen und nachfragen."
        )
    },
    "en": {
        "app_title": "PC Analysis Tool V5",
        "header_title": "PC Analysis",
        "header_subtitle": "Where does the disk space go and what slows down the PC? This tool only displays information and changes nothing.",
        "btn_safety": "Safety Information",
        "btn_cleanup": "Safe Cleanup",
        "tab_storage": "Storage",
        "tab_autostart": "Startup",
        "tab_folders": "Largest Folders & Files",
        "btn_refresh": "Refresh",
        "btn_drive_opt": "Open Windows Drive Optimizer",
        "lbl_trim_checking": "Checking TRIM setting…",
        "lbl_trim_result": "Windows TRIM setting: {}. This is a Windows setting, not a direct drive function test.",
        "lbl_storage_warn": "Warning when less than {}% is free. Click 'Analyze' to see what uses space. Drive optimization only opens the Windows tool.",
        "lbl_no_drives": "No drives found.",
        "lbl_drives_summary": "{} {} · {} total · {} used · {} free",
        "lbl_drives_low": " · ⚠ {} with low space",
        "unit_drive": "Drive",
        "unit_drives": "Drives",
        "chip_ok": "OK",
        "chip_warn": "Low Space",
        "chip_crit": "Almost Full",
        "lbl_free_of": "{} free of {}",
        "lbl_used_free": "{} used · {}% free",
        "lbl_type": "Type:",
        "lbl_model": "Model:",
        "lbl_interface": "Interface:",
        "lbl_health": "Reported Status:",
        "lbl_temp_wear": "Temperature/Wear: not provided by Windows",
        "lbl_temperature": "Temperature",
        "lbl_wear": "Wear",
        "health_good": "Good",
        "health_warning": "Warning",
        "health_critical": "Critical",
        "health_unknown": "Unknown",
        "btn_analyze": "Analyze Storage",
        "btn_choose": "Choose Folder & Analyze…",
        "btn_up": "↑ One Level Up",
        "btn_cancel": "Cancel",
        "btn_export": "Export…",
        "lbl_folder_not_chosen": "No folder chosen yet",
        "lbl_folder_hint": "Choose a folder or drive to start analysis. Click a column header to sort, right-click opens a menu.",
        "lbl_no_analysis": "No analysis yet.\n\nChoose a folder above or start a drive analysis in the 'Storage' tab.",
        "lbl_analyzing": "Analyzing… This may take a while for large folders.",
        "lbl_canceling": "Canceling...",
        "lbl_analysis_aborted": "Analysis aborted.",
        "lbl_analysis_failed": "Analysis failed.",
        "lbl_finished": "Done. {} in {} files.",
        "lbl_not_readable": " {} entries not readable (missing permissions), sum may be slightly low.",
        "lbl_active_entries": "{} active entries of {}",
        "lbl_missing_files": " · {} with missing file (red, often remnants of uninstalled software)",
        "lbl_autostart_win_only": "Autostart display only works on Windows.",
        "lbl_autostart_only_display": "Display only: Use Task Manager ('Startup' tab) to disable. Disabling is reversible, deleting is not: Do not remove anything from the registry. Hints are guidelines, not guarantees.",
        "btn_taskmgr": "Open Task Manager",
        "btn_startup_settings": "Windows Startup Settings",
        "error_title": "Error",
        "error_taskmgr_failed": "Task Manager could not be started:\n{}",
        "error_settings_failed": "Settings could not be opened:\n{}",
        "error_drive_opt_failed": "Windows Drive Optimizer could not be opened:\n{}\n\nTip: Start PC Analysis Tool once as Administrator (right-click .py file or icon → 'Run as administrator'), then it works without UAC prompt.",
        "error_not_available": "Not available",
        "error_drive_opt_win_only": "Windows Drive Optimizer is only available on Windows.",
        "safety_title": "Important Safety Information",
        "safety_warning_title": "Caution: Critical Area",
        "safety_warning_text": "{}\n\nYou may view the area. But do not delete, move or change anything there if you don't know exactly what it is. Windows or programs may not work properly afterwards.\n\nOpen in Explorer anyway?",
        "lbl_read_before_delete": "⚠  Please read briefly before deleting anything",
        "lbl_hide_startup_notice": "Do not show this notice automatically at startup",
        "btn_understood": "Understood",
        "error_export_failed": "Export failed:\n{}",
        "error_exported": "Exported",
        "error_exported_text": "Results saved:\n{}",
        "error_folder_not_found": "Folder not found:\n{}",
        "error_unexpected": "Unexpected Error",
        "col_program": "Program",
        "col_status": "Status",
        "col_file": "File",
        "col_hint": "Hint",
        "col_command": "Command / Path",
        "col_source": "Source",
        "status_active": "Active",
        "status_disabled": "Disabled",
        "status_present": "present",
        "status_missing": "missing",
        "hint_missing_file": "Program file missing – usually remnant of uninstallation",
        "hint_security": "⚠ Security software – better keep active",
        "hint_windows": "⚠ Windows component – do not change",
        "hint_driver": "⚠ Driver/Hardware tool – keep active when in doubt",
        "col_folder": "Folder",
        "col_size": "Size",
        "col_share": "Share",
        "col_newest": "Newest Change",
        "col_age": "Age",
        "col_changed": "Changed",
        "col_type": "File Type",
        "tab_folder": "Folders",
        "tab_files": "Largest Files",
        "tab_types": "File Types",
        "lbl_files_in_root": "(Files directly in this folder)",
        "export_csv": "CSV (Excel)",
        "export_txt": "Text Report",
        "export_default_name": "PC-Analysis_{}_{}.csv",
        "today": "today",
        "days": "{} days",
        "months": "{} months",
        "years": "{} years",
        "no_extension": "(no extension)",
        "ssd": "SSD",
        "hdd": "Hard Disk (HDD)",
        "scm": "Storage (SCM)",
        "nvme": "NVMe",
        "sata": "SATA",
        "usb": "USB",
        "sas": "SAS",
        "ata": "ATA",
        "raid": "RAID",
        "trim_enabled": "Enabled",
        "trim_disabled": "Disabled",
        "trim_not_available": "Not available",
        "open_in_explorer": "Open in Explorer",
        "show_in_explorer": "Show in Explorer",
        "copy_path": "Copy Path",
        "analyze_here": "Analyze this folder (go inside)",
        "lbl_main_hint": "⚠  Do not delete anything you don't know, especially nothing in Windows, program folders or registry. Critical areas are marked with ⛔ / ⚠.",
        "root_special_sysvol": "⛔ Windows recovery points – do not delete",
        "root_special_recovery": "⛔ Windows Recovery – do not delete",
        "root_special_winre": "⛔ Windows Update Recovery – do not delete",
        "root_special_boot": "⛔ Windows boot files – do not delete",
        "root_special_efi": "⛔ Windows boot files – do not delete",
        "root_special_bootmgr": "⛔ Windows boot file – do not delete",
        "root_special_pagefile": "⛔ Windows page file – do not delete",
        "root_special_swapfile": "⛔ Windows system file – do not delete",
        "root_special_hiberfil": "⛔ Hibernate file – do not delete",
        "root_special_recycle": "⚠ Recycle Bin – use 'Empty Recycle Bin'",
        "root_special_winold": "⚠ Old Windows version – only remove via Windows Disk Cleanup",
        "root_special_winbt": "⚠ Update remnants – only remove via Windows Disk Cleanup",
        "root_special_configmsi": "⚠ Installer remnants – do not delete manually",
        "loc_windir": "⛔ Windows System – do not delete",
        "loc_progfiles": "⚠ Program files – only remove via 'Uninstall apps'",
        "loc_progdata": "⚠ Program data – do not delete manually",
        "loc_appdata": "⚠ App data – do not delete manually",
        "safety_notice": (
            "• This tool only displays and deletes nothing. However, via Explorer (right-click → Open) you get "
            "direct access to files, so please be careful.\n\n"
            "• Never delete files or folders in C:\\Windows, Program Files, ProgramData or AppData and no "
            "system files like pagefile.sys or hiberfil.sys. Windows or programs may not start afterwards.\n\n"
            "• Large files are not automatically unnecessary. Only delete what you know yourself, for example your own "
            "downloads, old videos or installation files.\n\n"
            "• Never remove programs by deleting their folder: Settings → Apps (Installed Apps or "
            "Apps & Features) → Uninstall.\n\n"
            "• Autostart: Only disable entries (this is reversible), do not delete. Better keep security software and "
            "driver tools active.\n\n"
            "• Never change Windows registry by hand and do not use 'registry cleaners'. This tool only reads the "
            "registry.\n\n"
            "• Windows Storage Sense is suitable for safe cleanup (button 'Safe Cleanup'). "
            "Create a backup or restore point before major cleanup actions.\n\n"
            "• When in doubt: better not delete anything and ask."
        )
    }
}

def t(key):
    """Gibt den übersetzten Text für die erkannte Sprache zurück."""
    return TEXTS[LANG].get(key, TEXTS["de"].get(key, key))

# ----------------------------------------------------------------------------
# Konstanten & Formatierung
# ----------------------------------------------------------------------------
LOW_SPACE_PERCENT = 15
TOP_FOLDERS = 200
TOP_FILES = 25
TOP_TYPES = 40
SCAN_WORKERS = 4
REPARSE_POINT = 0x400

class Cancelled(Exception):
    pass

def fmt_int(n):
    return f"{n:,}".replace(",", ".")

def fmt_size(num):
    num = float(num)
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if num < 1024 or unit == "TB":
            return f"{num:.0f} {unit}" if unit == "B" else f"{num:.1f} {unit}".replace(".", ",")
        num /= 1024

def fmt_pct(value):
    return f"{value:.1f} %".replace(".", ",")

def fmt_date(ts):
    try:
        return time.strftime("%d.%m.%Y %H:%M", time.localtime(ts)) if ts > 0 else "-"
    except (OSError, OverflowError, ValueError):
        return "-"

def age_seconds(ts):
    return time.time() - ts if ts > 0 else -1

def fmt_age(ts):
    if ts <= 0:
        return "-"
    days = (time.time() - ts) / 86400
    if days < 1:
        return t("today")
    if days < 45:
        return t("days").format(int(days))
    if days < 700:
        return t("months").format(int(days / 30.4))
    return t("years").format(days / 365.25).replace(".", ",")

# ----------------------------------------------------------------------------
# System-Helfer
# ----------------------------------------------------------------------------
def list_drives():
    if os.name == "nt":
        import string
        return [f"{c}:\\" for c in string.ascii_uppercase if os.path.exists(f"{c}:\\")]
    return ["/"]

def read_drive_details():
    if os.name != "nt":
        return {"_trim": t("trim_not_available")}

    script = r"""
$ErrorActionPreference = 'SilentlyContinue'
$result = @()
$letters = Get-Volume | Where-Object { $_.DriveLetter } | Select-Object -ExpandProperty DriveLetter
foreach ($letter in $letters) {
    $disk = Get-Partition -DriveLetter $letter | Get-Disk | Select-Object -First 1
    if (-not $disk) { continue }
    $physical = Get-PhysicalDisk | Where-Object { [string]$_.DeviceId -eq [string]$disk.Number } | Select-Object -First 1
    $media = if ($physical -and $physical.MediaType -and $physical.MediaType -ne 'Unspecified') { [string]$physical.MediaType } else { 'Unknown' }
    $model = [string]$disk.FriendlyName
    $bus = [string]$disk.BusType
    $health = if ($physical) { [string]$physical.HealthStatus } else { [string]$disk.HealthStatus }
    $temperature = $null
    $wear = $null
    if ($physical) {
        $counter = Get-StorageReliabilityCounter -PhysicalDisk $physical
        if ($counter -and $counter.Temperature -gt 0) { $temperature = [int]$counter.Temperature }
        if ($counter -and $counter.Wear -ge 0) { $wear = [int]$counter.Wear }
    }
    $result += [pscustomobject]@{
        Letter = [string]$letter
        Model = $model
        Bus = $bus
        Media = $media
        Health = $health
        Temperature = $temperature
        Wear = $wear
    }
}
$result | ConvertTo-Json -Depth 3 -Compress
"""
    try:
        completed = subprocess.run(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", script],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=18,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)
        )
        raw = completed.stdout.strip()
        data = json.loads(raw) if raw else []
        if isinstance(data, dict):
            data = [data]
        details = {str(item.get("Letter", "")).upper(): item for item in data if item.get("Letter")}
    except (OSError, subprocess.SubprocessError, json.JSONDecodeError, TypeError, ValueError):
        details = {}

    trim = t("trim_not_available")
    try:
        completed = subprocess.run(
            ["fsutil.exe", "behavior", "query", "DisableDeleteNotify"],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=8,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)
        )
        output = (completed.stdout + "\n" + completed.stderr)
        values = re.findall(r"DisableDeleteNotify\s*=\s*(\d+)", output, flags=re.I)
        if values:
            trim = t("trim_enabled") if all(value == "0" for value in values) else t("trim_disabled")
    except (OSError, subprocess.SubprocessError):
        pass
    details["_trim"] = trim
    return details

def open_folder(path):
    try:
        if os.name == "nt":
            os.startfile(path)
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception as e:
        messagebox.showerror(t("error_title"), t("error_settings_failed").format(e))

def reveal_file(path):
    try:
        if os.name == "nt":
            subprocess.Popen(f'explorer /select,"{os.path.normpath(path)}"')
        else:
            subprocess.Popen(["xdg-open", os.path.dirname(path)])
    except Exception as e:
        messagebox.showerror(t("error_title"), t("error_settings_failed").format(e))

def is_link(st):
    return stat.S_ISLNK(st.st_mode) or bool(getattr(st, "st_file_attributes", 0) & REPARSE_POINT)

# ----------------------------------------------------------------------------
# Sicherheitshinweise
# ----------------------------------------------------------------------------
LEVEL_SYSTEM, LEVEL_PROGRAM, LEVEL_USERDATA = "system", "program", "userdata"

ROOT_SPECIAL = {
    "system volume information": (LEVEL_SYSTEM, t("root_special_sysvol")),
    "recovery": (LEVEL_SYSTEM, t("root_special_recovery")),
    "$winreagent": (LEVEL_SYSTEM, t("root_special_winre")),
    "boot": (LEVEL_SYSTEM, t("root_special_boot")),
    "efi": (LEVEL_SYSTEM, t("root_special_efi")),
    "bootmgr": (LEVEL_SYSTEM, t("root_special_bootmgr")),
    "pagefile.sys": (LEVEL_SYSTEM, t("root_special_pagefile")),
    "swapfile.sys": (LEVEL_SYSTEM, t("root_special_swapfile")),
    "hiberfil.sys": (LEVEL_SYSTEM, t("root_special_hiberfil")),
    "$recycle.bin": (LEVEL_PROGRAM, t("root_special_recycle")),
    "windows.old": (LEVEL_PROGRAM, t("root_special_winold")),
    "$windows.~bt": (LEVEL_PROGRAM, t("root_special_winbt")),
    "config.msi": (LEVEL_PROGRAM, t("root_special_configmsi")),
}

def protected_locations():
    result = []
    def add(var, level, text):
        value = os.environ.get(var)
        if value:
            result.append((os.path.normcase(os.path.abspath(value)), level, text))
    add("WINDIR", LEVEL_SYSTEM, t("loc_windir"))
    add("ProgramFiles", LEVEL_PROGRAM, t("loc_progfiles"))
    add("ProgramFiles(x86)", LEVEL_PROGRAM, t("loc_progfiles"))
    add("ProgramData", LEVEL_PROGRAM, t("loc_progdata"))
    add("LOCALAPPDATA", LEVEL_USERDATA, t("loc_appdata"))
    add("APPDATA", LEVEL_USERDATA, t("loc_appdata"))
    return result

def classify_path(path, locations=None):
    p = os.path.normcase(os.path.abspath(path))
    parent, name = os.path.split(p)
    if os.path.dirname(parent) == parent and name in ROOT_SPECIAL:
        return ROOT_SPECIAL[name]
    for base, level, text in (protected_locations() if locations is None else locations):
        try:
            if os.path.commonpath([p, base]) == base:
                return level, text
        except ValueError:
            continue
    return None

SECURITY_RE = re.compile(r"securityhealth|defender|antivir|kaspersky|mcafee|norton|bitdefender|malwarebytes|"
                         r"sophos|avira|\bavast|\bavg\b|\beset\b", re.I)
HARDWARE_RE = re.compile(r"realtek|rtkaud|nvidia|nvcpl|radeon|synaptics|touchpad|\bamd\b|\bintel\b|bluetooth", re.I)

def autostart_hint(name, command, target):
    if target == t("status_missing"):
        return t("hint_missing_file")
    text = f"{name} {command}"
    if SECURITY_RE.search(text):
        return t("hint_security")
    windir = os.environ.get("WINDIR")
    if windir and os.path.normcase(command.strip().strip('"')).startswith(os.path.normcase(windir)):
        return t("hint_windows")
    if HARDWARE_RE.search(text):
        return t("hint_driver")
    return ""

def settings_file():
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    return os.path.join(base, "PC-Analyse-Tool", "settings.json")

def load_settings():
    try:
        with open(settings_file(), encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}

def save_settings(data):
    try:
        path = settings_file()
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f)
    except OSError:
        pass

# ----------------------------------------------------------------------------
# Ordner-Analyse
# ----------------------------------------------------------------------------
class Counter:
    def __init__(self):
        self.lock = threading.Lock()
        self.files = 0
        self.bytes = 0
        self.errors = 0

    def add(self, files=0, nbytes=0, errors=0):
        with self.lock:
            self.files += files
            self.bytes += nbytes
            self.errors += errors

def push_top(heap, size, path, mtime, n=TOP_FILES):
    item = (size, path, mtime)
    if len(heap) < n:
        heapq.heappush(heap, item)
    elif size > heap[0][0]:
        heapq.heapreplace(heap, item)

def file_ext(name):
    return os.path.splitext(name)[1].lower() or t("no_extension")

def dir_size(path, cancel, counter):
    total, newest = 0, 0.0
    top = []
    ext = defaultdict(int)
    pending_files = pending_bytes = 0
    stack = [path]
    while stack:
        if cancel.is_set():
            break
        current = stack.pop()
        try:
            with os.scandir(current) as it:
                for entry in it:
                    try:
                        st = entry.stat(follow_symlinks=False)
                    except OSError:
                        counter.add(errors=1)
                        continue
                    if is_link(st):
                        continue
                    if stat.S_ISDIR(st.st_mode):
                        stack.append(entry.path)
                    elif stat.S_ISREG(st.st_mode):
                        size = st.st_size
                        total += size
                        if st.st_mtime > newest:
                            newest = st.st_mtime
                        push_top(top, size, entry.path, st.st_mtime)
                        ext[file_ext(entry.name)] += size
                        pending_files += 1
                        pending_bytes += size
                        if pending_files >= 256:
                            counter.add(pending_files, pending_bytes)
                            pending_files = pending_bytes = 0
        except OSError:
            counter.add(errors=1)
    counter.add(pending_files, pending_bytes)
    return total, top, newest, dict(ext)

def scan_folder(folder, cancel, post):
    counter = Counter()
    with os.scandir(folder) as it:
        entries = list(it)

    total_units = max(len(entries), 1)
    post("progress", 0, total_units)
    post("status", f"Reading main folder ({fmt_int(len(entries))} entries)..." if LANG == "en" else f"Lese Hauptordner ({fmt_int(len(entries))} Einträge)...")

    subdirs = []
    root_size, root_newest = 0, 0.0
    top = []
    ext = defaultdict(int)
    done_files = 0
    pending_files = pending_bytes = 0

    for index, entry in enumerate(entries, 1):
        if cancel.is_set():
            raise Cancelled
        try:
            st = entry.stat(follow_symlinks=False)
        except OSError:
            counter.add(errors=1)
            done_files += 1
            continue
        if stat.S_ISDIR(st.st_mode) and not is_link(st):
            subdirs.append((entry.name, entry.path))
        else:
            done_files += 1
            if stat.S_ISREG(st.st_mode) and not is_link(st):
                root_size += st.st_size
                root_newest = max(root_newest, st.st_mtime)
                push_top(top, st.st_size, entry.path, st.st_mtime)
                ext[file_ext(entry.name)] += st.st_size
                pending_files += 1
                pending_bytes += st.st_size
        if index % 1000 == 0:
            counter.add(pending_files, pending_bytes)
            pending_files = pending_bytes = 0
            post("progress", done_files, total_units)
            status_text = f"Reading main folder... {fmt_int(index)}/{fmt_int(len(entries))} entries" if LANG == "en" else f"Lese Hauptordner... {fmt_int(index)}/{fmt_int(len(entries))} Einträge"
            post("status", status_text)
    counter.add(pending_files, pending_bytes)
    post("progress", done_files, total_units)

    folders = []
    finished = 0
    with ThreadPoolExecutor(max_workers=SCAN_WORKERS) as ex:
        futures = {ex.submit(dir_size, path, cancel, counter): (name, path) for name, path in subdirs}
        pending = set(futures)
        try:
            while pending:
                if cancel.is_set():
                    raise Cancelled
                done_set, pending = wait(pending, timeout=0.25, return_when=FIRST_COMPLETED)
                for fut in done_set:
                    name, path = futures[fut]
                    size, sub_top, newest, sub_ext = fut.result()
                    folders.append((name, path, size, newest))
                    for s, p, m in sub_top:
                        push_top(top, s, p, m)
                    for k, v in sub_ext.items():
                        ext[k] += v
                    finished += 1
                post("progress", done_files + finished, total_units)
                note = f" · {fmt_int(counter.errors)} not readable" if LANG == "en" else f" · {fmt_int(counter.errors)} nicht lesbar" if counter.errors else ""
                status_text = f"Analyzing... {fmt_int(counter.files)} files, {fmt_size(counter.bytes)} read · Folders {finished}/{len(subdirs)}{note}" if LANG == "en" else f"Analysiere... {fmt_int(counter.files)} Dateien, {fmt_size(counter.bytes)} gelesen · Ordner {finished}/{len(subdirs)}{note}"
                post("status", status_text)
        finally:
            for f in futures:
                f.cancel()

    folders.sort(key=lambda r: -r[2])
    return {
        "folder": folder,
        "folders": folders,
        "files": sorted(top, reverse=True),
        "ext": sorted(ext.items(), key=lambda kv: -kv[1]),
        "root_files": root_size,
        "root_newest": root_newest,
        "total": sum(f[2] for f in folders) + root_size,
        "file_count": counter.files,
        "errors": counter.errors,
        "time": time.time(),
    }

# ----------------------------------------------------------------------------
# Export
# ----------------------------------------------------------------------------
def build_csv_rows(r):
    total = r["total"] or 1
    rows = [[t("col_type"), "Path", "Size (Bytes)", t("col_size"), t("col_share"), t("col_newest") + " / " + t("col_changed")]]
    for name, path, size, newest in r["folders"][:TOP_FOLDERS]:
        rows.append([t("col_folder"), path, size, fmt_size(size), fmt_pct(size / total * 100), fmt_date(newest)])
    if r["root_files"]:
        rows.append([t("lbl_files_in_root"), r["folder"], r["root_files"], fmt_size(r["root_files"]),
                     fmt_pct(r["root_files"] / total * 100), fmt_date(r["root_newest"])])
    for size, path, mtime in r["files"]:
        rows.append(["File", path, size, fmt_size(size), fmt_pct(size / total * 100), fmt_date(mtime)])
    for ext, size in r["ext"][:TOP_TYPES]:
        rows.append([t("col_type"), ext, size, fmt_size(size), fmt_pct(size / total * 100), ""])
    return rows

def build_text_report(r):
    total = r["total"]
    def pct(size):
        return fmt_pct(size / total * 100) if total else "-"
    lines = [
        "PC Analysis - Report" if LANG == "en" else "PC-Analyse - Bericht",
        "=" * 78,
        f"Folder:       {r['folder']}" if LANG == "en" else f"Ordner:       {r['folder']}",
        f"Created:      {fmt_date(r['time'])}" if LANG == "en" else f"Erstellt am:  {fmt_date(r['time'])}",
        f"Total size:   {fmt_size(total)} in {fmt_int(r['file_count'])} files" if LANG == "en" else f"Gesamtgröße:  {fmt_size(total)} in {fmt_int(r['file_count'])} Dateien",
    ]
    if r["errors"]:
        lines.append(f"Not readable: {fmt_int(r['errors'])} entries (missing permissions); sums may be too low" if LANG == "en" else f"Nicht lesbar: {fmt_int(r['errors'])} Einträge (fehlende Rechte); Summen können zu klein sein")
    shown = r["folders"][:TOP_FOLDERS]
    lines += ["", f"LARGEST FOLDERS ({len(shown)})" if LANG == "en" else f"GRÖSSTE ORDNER ({len(shown)})", "-" * 78,
              f"{'Size':>10}  {'Share':>8}  {'Newest Change':<16}  Folder" if LANG == "en" else f"{'Größe':>10}  {'Anteil':>8}  {'Neueste Änderung':<16}  Ordner"]
    for name, path, size, newest in shown:
        lines.append(f"{fmt_size(size):>10}  {pct(size):>8}  {fmt_date(newest):<16}  {path}")
    if r["root_files"]:
        lines.append(f"{fmt_size(r['root_files']):>10}  {pct(r['root_files']):>8}  "
                     f"{fmt_date(r['root_newest']):<16}  (Files directly in this folder)" if LANG == "en" else f"{fmt_date(r['root_newest']):<16}  (Dateien direkt im Hauptordner)")
    lines += ["", f"LARGEST FILES ({len(r['files'])})" if LANG == "en" else f"GRÖSSTE DATEIEN ({len(r['files'])})", "-" * 78,
              f"{'Size':>10}  {'Changed':<16}  {'Age':>10}  File" if LANG == "en" else f"{'Größe':>10}  {'Geändert':<16}  {'Alter':>10}  Datei"]
    for size, path, mtime in r["files"]:
        lines.append(f"{fmt_size(size):>10}  {fmt_date(mtime):<16}  {fmt_age(mtime):>10}  {path}")
    types = r["ext"][:TOP_TYPES]
    lines += ["", f"FILE TYPES ({len(types)})" if LANG == "en" else f"DATEITYPEN ({len(types)})", "-" * 78, f"{'Size':>10}  {'Share':>8}  Type" if LANG == "en" else f"{'Größe':>10}  {'Anteil':>8}  Typ"]
    for ext, size in types:
        lines.append(f"{fmt_size(size):>10}  {pct(size):>8}  {ext}")
    return "\n".join(lines) + "\n"

# ----------------------------------------------------------------------------
# Autostart auslesen
# ----------------------------------------------------------------------------
RUN_BASE = r"Software\Microsoft\Windows\CurrentVersion"
APPROVED_BASE = RUN_BASE + r"\Explorer\StartupApproved"
TARGET_RE = re.compile(r'^\s*(?:"([^"]+)"|(.+?\.(?:exe|bat|cmd|com|lnk|msc|cpl|scr)))(?=\s|$)', re.I)

def target_status(command):
    m = TARGET_RE.match(command or "")
    if not m:
        return "?"
    path = os.path.expandvars((m.group(1) or m.group(2)).strip())
    return t("status_present") if (os.path.exists(path) or shutil.which(path)) else t("status_missing")

def approved_state(hive, approved_sub, name):
    try:
        with winreg.OpenKey(hive, f"{APPROVED_BASE}\\{approved_sub}") as key:
            data, _ = winreg.QueryValueEx(key, name)
        if isinstance(data, (bytes, bytearray)) and data and (data[0] & 1):
            return t("status_disabled")
    except OSError:
        pass
    return t("status_active")

def read_autostart():
    items = []
    if winreg is None:
        return items
    run_sources = [
        (winreg.HKEY_CURRENT_USER, RUN_BASE + r"\Run", "Run", "User (Registry)" if LANG == "en" else "Benutzer (Registry)"),
        (winreg.HKEY_LOCAL_MACHINE, RUN_BASE + r"\Run", "Run", "All Users (Registry)" if LANG == "en" else "Alle Benutzer (Registry)"),
        (winreg.HKEY_LOCAL_MACHINE, r"Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Run",
         "Run32", "All Users, 32-Bit (Registry)" if LANG == "en" else "Alle Benutzer, 32-Bit (Registry)"),
    ]
    for hive, path, approved_sub, label in run_sources:
        try:
            with winreg.OpenKey(hive, path) as key:
                i = 0
                while True:
                    try:
                        name, value, _ = winreg.EnumValue(key, i)
                    except OSError:
                        break
                    i += 1
                    items.append((name, approved_state(hive, approved_sub, name), str(value), label))
        except OSError:
            continue
    folder_sources = [
        (os.path.join(os.environ.get("APPDATA", ""), r"Microsoft\Windows\Start Menu\Programs\Startup"),
         winreg.HKEY_CURRENT_USER, "Startup Folder (User)" if LANG == "en" else "Autostart-Ordner (Benutzer)"),
        (os.path.join(os.environ.get("PROGRAMDATA", ""), r"Microsoft\Windows\Start Menu\Programs\StartUp"),
         winreg.HKEY_LOCAL_MACHINE, "Startup Folder (All Users)" if LANG == "en" else "Autostart-Ordner (Alle Benutzer)"),
    ]
    for folder, hive, label in folder_sources:
        if not folder or not os.path.isdir(folder):
            continue
        try:
            for fname in os.listdir(folder):
                if fname.lower() == "desktop.ini":
                    continue
                items.append((fname, approved_state(hive, "StartupFolder", fname),
                              os.path.join(folder, fname), label))
        except OSError:
            continue
    items.sort(key=lambda r: (r[1] != t("status_active"), r[0].lower()))
    return items

# ----------------------------------------------------------------------------
# Design (Farben, Schriften, Stile)
# ----------------------------------------------------------------------------
BG = "#f4f5f7"
CARD = "#ffffff"
BORDER = "#d8dce2"
TEXT = "#1f2937"
MUTED = "#6b7280"
ACCENT = "#2563eb"
ACCENT_DARK = "#1d4ed8"
OK_COLOR, WARN_COLOR, CRIT_COLOR = "#16a34a", "#d97706", "#dc2626"
AMBER_BG, AMBER_BORDER, AMBER_TEXT = "#fff4d6", "#f1d99a", "#7a4b00"
ZEBRA = "#f6f7f9"

FONT = ("Segoe UI", 12)
FONT_BOLD = ("Segoe UI", 12, "bold")
FONT_SMALL = ("Segoe UI", 11)
FONT_TITLE = ("Segoe UI", 22, "bold")
FONT_BIG = ("Segoe UI", 24, "bold")

CARD_COLUMNS = 2
CHIPS = {
    "Ok": (t("chip_ok"), "#dcfce7", "#166534"),
    "Warn": (t("chip_warn"), "#fef3c7", "#92400e"),
    "Crit": (t("chip_crit"), "#fee2e2", "#991b1b"),
}

def apply_theme(root):
    root.configure(bg=BG)
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass
    style.configure(".", background=BG, foreground=TEXT, font=FONT)
    style.configure("TFrame", background=BG)
    style.configure("TLabel", background=BG, foreground=TEXT)
    style.configure("TButton", padding=(16, 8), font=FONT, background="#e5e7eb", foreground=TEXT,
                    borderwidth=0, focusthickness=0, relief="flat")
    style.map("TButton", background=[("disabled", "#eceef1"), ("active", "#d9dde3")],
              foreground=[("disabled", "#9ca3af")])
    style.configure("Primary.TButton", background=ACCENT, foreground="white", font=FONT_BOLD)
    style.map("Primary.TButton", background=[("disabled", "#93b4f0"), ("active", ACCENT_DARK)],
              foreground=[("disabled", "white")])
    style.configure("Treeview", background=CARD, fieldbackground=CARD, foreground=TEXT, rowheight=32,
                    borderwidth=0, font=FONT)
    style.map("Treeview", background=[("selected", "#dbe7fd")], foreground=[("selected", TEXT)])
    style.configure("Treeview.Heading", background="#eef0f3", foreground=TEXT, font=FONT_BOLD,
                    padding=(12, 10), relief="flat")
    style.map("Treeview.Heading", background=[("active", "#e1e5ea")])
    for name, color in (("Ok", OK_COLOR), ("Warn", WARN_COLOR), ("Crit", CRIT_COLOR), ("Accent", ACCENT)):
        style.configure(f"{name}.Horizontal.TProgressbar", troughcolor="#e5e7eb", background=color,
                        lightcolor=color, darkcolor=color, bordercolor="#e5e7eb", thickness=14)

def make_label(parent, text="", font=FONT, fg=TEXT, bg=BG, **kw):
    return tk.Label(parent, text=text, font=font, fg=fg, bg=bg, **kw)

def make_card(parent):
    return tk.Frame(parent, bg=CARD, highlightthickness=1, highlightbackground=BORDER, highlightcolor=BORDER)

def auto_wrap(label, container, margin=30):
    container.bind("<Configure>", lambda e: label.config(wraplength=max(e.width - margin, 200)), add="+")

# ----------------------------------------------------------------------------
# Tabelle mit Spaltensortierung
# ----------------------------------------------------------------------------
class SortableTree(ttk.Treeview):
    def __init__(self, parent, spec, numeric=(), height=10):
        ids = [c[0] for c in spec if c[0] != "#0"]
        super().__init__(parent, columns=ids, height=height)
        self.titles = {c[0]: c[1] for c in spec}
        self.numeric = set(numeric)
        self.keys = {}
        self.paths = {}
        self.base_tags = {}
        self.sort_col = None
        self.sort_rev = False
        for cid, title, width, anchor in spec:
            self.heading(cid, text=title, anchor=anchor, command=lambda c=cid: self.sort_by(c))
            self.column(cid, width=width, anchor=anchor)

    def add(self, text, values, keys=None, path=None, tags=()):
        index = len(self.get_children(""))
        tags = tuple(tags)
        iid = self.insert("", "end", text=text, values=values, tags=tags + (("odd",) if index % 2 else ()))
        self.base_tags[iid] = tags
        self.keys[iid] = keys or {}
        if path:
            self.paths[iid] = path
        return iid

    def clear(self):
        self.delete(*self.get_children())
        self.keys.clear()
        self.paths.clear()
        self.base_tags.clear()

    def restripe(self):
        for index, iid in enumerate(self.get_children("")):
            self.item(iid, tags=self.base_tags.get(iid, ()) + (("odd",) if index % 2 else ()))

    def _key(self, iid, col):
        value = self.keys.get(iid, {}).get(col)
        if value is None:
            value = self.item(iid, "text") if col == "#0" else self.set(iid, col)
            return (1, str(value).lower())
        if isinstance(value, str):
            return (1, value.lower())
        return (0, value)

    def sort_by(self, col, reverse=None):
        if reverse is None:
            reverse = (not self.sort_rev) if self.sort_col == col else (col in self.numeric)
        items = sorted(self.get_children(""), key=lambda i: self._key(i, col), reverse=reverse)
        for index, iid in enumerate(items):
            self.move(iid, "", index)
        self.sort_col, self.sort_rev = col, reverse
        for cid, title in self.titles.items():
            arrow = (" ▼" if reverse else " ▲") if cid == col else ""
            self.heading(cid, text=title + arrow)
        self.restripe()

def make_tree(parent, spec, numeric=(), height=10):
    card = make_card(parent)
    card.pack(fill=tk.BOTH, expand=True)
    tree = SortableTree(card, spec, numeric, height)
    tree.tag_configure("odd", background=ZEBRA)
    scroll = ttk.Scrollbar(card, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=scroll.set)
    scroll.pack(side=tk.RIGHT, fill=tk.Y)
    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    return tree

class SafetyDialog(tk.Toplevel):
    def __init__(self, parent, hide_next):
        super().__init__(parent, bg=BG)
        self.title(t("safety_title"))
        self.resizable(False, False)
        self.transient(parent)
        self.hide_next = hide_next
        self.var = tk.IntVar(value=1 if hide_next else 0)
        tk.Label(self, text=t("lbl_read_before_delete"), font=("Segoe UI", 14, "bold"),
                 fg=WARN_COLOR, bg=BG).pack(padx=24, pady=(18, 10), anchor="w")
        box = make_card(self)
        box.pack(padx=24, fill=tk.X)
        tk.Label(box, text=t("safety_notice"), justify="left", wraplength=620, font=FONT, bg=CARD,
                 fg=TEXT).pack(padx=18, pady=14, anchor="w")
        tk.Checkbutton(self, text=t("lbl_hide_startup_notice"), variable=self.var,
                       bg=BG, activebackground=BG, font=FONT_SMALL, anchor="w").pack(padx=24, pady=(12, 0), anchor="w")
        ttk.Button(self, text=t("btn_understood"), style="Primary.TButton", command=self.close).pack(pady=16)
        self.protocol("WM_DELETE_WINDOW", self.close)
        self.update_idletasks()
        self.geometry(f"+{parent.winfo_rootx() + 80}+{parent.winfo_rooty() + 40}")
        self.grab_set()

    def close(self):
        self.hide_next = bool(self.var.get())
        self.destroy()

# ----------------------------------------------------------------------------
# Anwendung
# ----------------------------------------------------------------------------
class AnalyseApp:
    def __init__(self, root):
        self.root = root
        self.root.title(t("app_title"))
        self.setup_window()
        apply_theme(root)

        self.q = queue.Queue()
        self.cancel = threading.Event()
        self.busy = False
        self.scan_path = ""
        self.result = None
        self.settings = load_settings()
        self.drive_details = {}
        self.current_tab = 0

        self.build_header()
        self.build_banner()
        self.build_custom_tabs()

        self.tab_space = tk.Frame(self.tab_container, bg=BG, padx=6, pady=14)
        self.tab_auto = tk.Frame(self.tab_container, bg=BG, padx=6, pady=14)
        self.tab_folders = tk.Frame(self.tab_container, bg=BG, padx=6, pady=14)

        self.show_tab(0)

        self.build_space_tab()
        self.build_autostart_tab()
        self.build_folders_tab()

        self.refresh_space()
        self.refresh_autostart()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.after(50, self.poll_queue)
        self.root.after(300, lambda: self.show_safety_notice(startup=True))

    def setup_window(self):
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        w, h = min(1400, int(sw * 0.9)), min(900, int(sh * 0.88))
        self.root.geometry(f"{w}x{h}+{(sw - w) // 2}+{max((sh - h) // 2 - 20, 0)}")
        self.root.minsize(min(1100, w), min(700, h))
        self.root.resizable(False, False)

    def build_header(self):
        header = tk.Frame(self.root, bg=BG)
        header.pack(fill=tk.X, padx=24, pady=(20, 4))
        left = tk.Frame(header, bg=BG)
        left.pack(side=tk.LEFT)
        make_label(left, t("header_title"), FONT_TITLE).pack(anchor="w")
        make_label(left, t("header_subtitle"), FONT_SMALL, MUTED).pack(anchor="w", pady=(2, 0))
        right = tk.Frame(header, bg=BG)
        right.pack(side=tk.RIGHT)
        ttk.Button(right, text=t("btn_safety"), command=self.show_safety_notice).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(right, text=t("btn_cleanup"), style="Primary.TButton",
                   command=self.open_storage_settings).pack(side=tk.LEFT)

    def build_banner(self):
        banner = tk.Frame(self.root, bg=AMBER_BG, highlightthickness=1, highlightbackground=AMBER_BORDER)
        banner.pack(fill=tk.X, padx=24, pady=(10, 0))
        label = tk.Label(banner, bg=AMBER_BG, fg=AMBER_TEXT, font=FONT_SMALL, anchor="w", justify="left",
                         text=t("lbl_main_hint"))
        label.pack(fill=tk.X, padx=14, pady=8)
        auto_wrap(label, banner, 40)

    def build_custom_tabs(self):
        tab_frame = tk.Frame(self.root, bg=BG)
        tab_frame.pack(fill=tk.X, padx=24, pady=(10, 0))

        self.tab_buttons = []
        tab_names = [t("tab_storage"), t("tab_autostart"), t("tab_folders")]

        for i, name in enumerate(tab_names):
            btn = tk.Button(tab_frame, text=name, font=FONT_BOLD, bg="#e5e7eb", fg=MUTED,
                           relief="flat", borderwidth=1, padx=20, pady=12,
                           command=lambda idx=i: self.show_tab(idx))
            btn.pack(side=tk.LEFT, padx=(0, 2))
            self.tab_buttons.append(btn)

        self.tab_container = tk.Frame(self.root, bg=BG)
        self.tab_container.pack(fill=tk.BOTH, expand=True, padx=24, pady=(0, 20))

    def show_tab(self, index):
        self.current_tab = index
        for i, btn in enumerate(self.tab_buttons):
            if i == index:
                btn.config(bg=CARD, fg=ACCENT, relief="sunken")
            else:
                btn.config(bg="#e5e7eb", fg=MUTED, relief="flat")

        self.tab_space.pack_forget()
        self.tab_auto.pack_forget()
        self.tab_folders.pack_forget()

        if index == 0:
            self.tab_space.pack(fill=tk.BOTH, expand=True)
        elif index == 1:
            self.tab_auto.pack(fill=tk.BOTH, expand=True)
        else:
            self.tab_folders.pack(fill=tk.BOTH, expand=True)

    def build_space_tab(self):
        top = tk.Frame(self.tab_space, bg=BG)
        top.pack(fill=tk.X)
        self.lbl_space_summary = make_label(top, "", FONT_BOLD)
        self.lbl_space_summary.pack(side=tk.LEFT)
        ttk.Button(top, text=t("btn_drive_opt"),
                   command=self.open_drive_optimization).pack(side=tk.RIGHT, padx=(8, 0))
        ttk.Button(top, text=t("btn_refresh"), command=self.refresh_space).pack(side=tk.RIGHT)
        self.lbl_trim = make_label(self.tab_space, t("lbl_trim_checking"), FONT_SMALL, MUTED)
        self.lbl_trim.pack(anchor="w", pady=(5, 0))
        make_label(self.tab_space, t("lbl_storage_warn").format(LOW_SPACE_PERCENT),
                   FONT_SMALL, MUTED).pack(anchor="w", pady=(3, 8))

        scroll_container = tk.Frame(self.tab_space, bg=BG)
        scroll_container.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(scroll_container, bg=BG, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(scroll_container, orient="vertical", command=self.canvas.yview)
        self.cards_frame = tk.Frame(self.canvas, bg=BG)

        self.cards_frame.bind("<Configure>", lambda e: self._update_scrollregion())
        self.cards_window = self.canvas.create_window((0, 0), window=self.cards_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.bind("<Configure>", self._resize_drive_cards)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        def _on_mousewheel(event):
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            self.root.update_idletasks()

        self.tab_space.bind("<Enter>", lambda e: self.root.bind_all("<MouseWheel>", _on_mousewheel))
        self.tab_space.bind("<Leave>", lambda e: self.root.unbind_all("<MouseWheel>"))

        for col in range(CARD_COLUMNS):
            self.cards_frame.columnconfigure(col, weight=1, uniform="cards")

    def _update_scrollregion(self):
        self.root.update_idletasks()
        bbox = self.canvas.bbox("all")
        if bbox:
            self.canvas.configure(scrollregion=bbox)

    def _resize_drive_cards(self, event):
        self.canvas.itemconfigure(self.cards_window, width=event.width)
        self.root.update_idletasks()

    @staticmethod
    def free_percent(usage):
        return usage.free / usage.total * 100 if usage.total else 0

    def refresh_space(self):
        for child in self.cards_frame.winfo_children():
            child.destroy()
        drives = []
        self.drive_details = read_drive_details()
        trim_state = self.drive_details.get("_trim", t("trim_not_available"))
        trim_color = OK_COLOR if trim_state == t("trim_enabled") else (WARN_COLOR if trim_state == t("trim_disabled") else MUTED)
        self.lbl_trim.config(
            text=t("lbl_trim_result").format(trim_state),
            fg=trim_color)
        for drive in list_drives():
            try:
                drives.append((drive, shutil.disk_usage(drive)))
            except OSError:
                continue
        if not drives:
            make_label(self.cards_frame, t("lbl_no_drives"), fg=CRIT_COLOR).pack(pady=20)
            self.lbl_space_summary.config(text="")
            return
        total = sum(u.total for _, u in drives)
        used = sum(u.used for _, u in drives)
        free = sum(u.free for _, u in drives)
        low = sum(1 for _, u in drives if self.free_percent(u) < LOW_SPACE_PERCENT)
        unit = t("unit_drive") if len(drives) == 1 else t("unit_drives")
        summary = t("lbl_drives_summary").format(len(drives), unit, fmt_size(total), fmt_size(used), fmt_size(free))
        if low:
            summary += t("lbl_drives_low").format(low)
        self.lbl_space_summary.config(text=summary)
        for index, (drive, usage) in enumerate(drives):
            card = self.drive_card(self.cards_frame, drive, usage,
                                   self.drive_details.get(drive[0].upper(), {}))
            card.grid(row=index // CARD_COLUMNS, column=index % CARD_COLUMNS, sticky="nsew", padx=8, pady=8)
        self._update_scrollregion()

    def drive_card(self, parent, drive, usage, details=None):
        details = details or {}
        free_pct = self.free_percent(usage)
        used_pct = 100 - free_pct
        level = "Crit" if free_pct < 5 else ("Warn" if free_pct < LOW_SPACE_PERCENT else "Ok")
        chip_text, chip_bg, chip_fg = CHIPS[level]
        card = make_card(parent)
        inner = tk.Frame(card, bg=CARD)
        inner.pack(fill=tk.BOTH, expand=True, padx=20, pady=16)
        row = tk.Frame(inner, bg=CARD)
        row.pack(fill=tk.X)
        name = drive.rstrip("\\/") or drive
        tk.Label(row, text=name, font=FONT_BIG, bg=CARD, fg=TEXT).pack(side=tk.LEFT)
        tk.Label(row, text=chip_text, font=FONT_SMALL, bg=chip_bg, fg=chip_fg, padx=10, pady=3).pack(side=tk.RIGHT)
        ttk.Progressbar(inner, style=f"{level}.Horizontal.TProgressbar", maximum=100,
                        value=used_pct).pack(fill=tk.X, pady=(12, 10))
        tk.Label(inner, text=t("lbl_free_of").format(fmt_size(usage.free), fmt_size(usage.total)), font=FONT_BOLD,
                 bg=CARD, fg=TEXT, anchor="w").pack(fill=tk.X)
        tk.Label(inner, text=t("lbl_used_free").format(fmt_size(usage.used), fmt_pct(free_pct)), font=FONT_SMALL,
                 bg=CARD, fg=MUTED, anchor="w").pack(fill=tk.X, pady=(2, 0))

        media = str(details.get("Media", "Unknown") or "Unknown")
        media_map = {"SSD": t("ssd"), "HDD": t("hdd"), "SCM": t("scm")}
        media_text = media_map.get(media.upper(), t("lbl_drive_type_unknown") if LANG == "en" else "Laufwerkstyp nicht erkennbar")
        model = str(details.get("Model", "") or (t("lbl_model_not_available") if LANG == "en" else "Modell nicht verfügbar"))
        bus_raw = str(details.get("Bus", "") or "Unknown")
        bus_map = {"NVME": t("nvme"), "SATA": t("sata"), "USB": t("usb"), "SAS": t("sas"), "ATA": t("ata"), "RAID": t("raid")}
        bus = bus_map.get(bus_raw.upper(), bus_raw)

        health_raw = str(details.get("Health", "") or "Unknown")
        health_map = {
            "HEALTHY": t("health_good"),
            "WARNING": t("health_warning"),
            "UNHEALTHY": t("health_critical"),
            "UNKNOWN": t("health_unknown")
        }
        health = health_map.get(health_raw.upper(), health_raw)

        info = tk.Frame(inner, bg=CARD)
        info.pack(fill=tk.X, pady=(10, 0))
        for label, value in ((t("lbl_type"), media_text), (t("lbl_model"), model), (t("lbl_interface"), bus),
                             (t("lbl_health"), health)):
            line = tk.Frame(info, bg=CARD)
            line.pack(fill=tk.X, pady=1)
            tk.Label(line, text=label, font=FONT_SMALL, bg=CARD, fg=MUTED, anchor="w", width=18).pack(side=tk.LEFT)
            tk.Label(line, text=value, font=FONT_SMALL, bg=CARD, fg=TEXT, anchor="w", justify="left",
                     wraplength=360).pack(side=tk.LEFT, fill=tk.X, expand=True)

        temperature = details.get("Temperature")
        wear = details.get("Wear")
        extra = []
        if temperature is not None:
            extra.append(f"{t('lbl_temperature')}: {temperature} °C")
        if wear is not None:
            extra.append(f"{t('lbl_wear')}: {wear} %")
        if extra:
            tk.Label(info, text=" · ".join(extra), font=FONT_SMALL, bg=CARD, fg=MUTED, anchor="w").pack(fill=tk.X, pady=(3, 0))
        elif os.name == "nt":
            tk.Label(info, text=t("lbl_temp_wear"), font=FONT_SMALL,
                     bg=CARD, fg=MUTED, anchor="w").pack(fill=tk.X, pady=(3, 0))

        ttk.Button(inner, text=t("btn_analyze"),
                   command=lambda d=drive: self.start_scan(d)).pack(anchor="e", pady=(12, 0))
        return card

    def open_drive_optimization(self):
        try:
            if os.name == "nt":
                os.startfile("dfrgui.exe")
            else:
                messagebox.showinfo(t("error_not_available"), t("error_drive_opt_win_only"))
        except OSError as e:
            messagebox.showerror(
                t("error_title"),
                t("error_drive_opt_failed").format(e)
            )

    def build_autostart_tab(self):
        top = tk.Frame(self.tab_auto, bg=BG)
        top.pack(fill=tk.X)
        self.lbl_auto_summary = make_label(top, "", FONT_BOLD)
        self.lbl_auto_summary.pack(side=tk.LEFT)
        ttk.Button(top, text=t("btn_refresh"), command=self.refresh_autostart).pack(side=tk.RIGHT)
        ttk.Button(top, text=t("btn_startup_settings"),
                   command=self.open_startup_settings).pack(side=tk.RIGHT, padx=(0, 8))
        ttk.Button(top, text=t("btn_taskmgr"), command=self.open_taskmgr).pack(side=tk.RIGHT, padx=(0, 8))
        self.lbl_auto = make_label(self.tab_auto, "", FONT_SMALL, MUTED, justify="left", anchor="w")
        self.lbl_auto.pack(fill=tk.X, pady=(6, 10))
        auto_wrap(self.lbl_auto, self.tab_auto, 20)
        self.tree_auto = make_tree(
            self.tab_auto,
            (("#0", t("col_program"), 190, "w"), ("status", t("col_status"), 95, "w"), ("datei", t("col_file"), 85, "w"),
             ("hinweis", t("col_hint"), 280, "w"), ("befehl", t("col_command"), 260, "w"), ("quelle", t("col_source"), 160, "w")))
        self.tree_auto.tag_configure("off", foreground="gray")
        self.tree_auto.tag_configure("missing", foreground="#b23b3b")
        self.tree_auto.tag_configure("warn", foreground="#b26a00")

    def refresh_autostart(self):
        self.tree_auto.clear()
        self.lbl_auto_summary.config(text="")
        if winreg is None:
            self.lbl_auto.config(text=t("lbl_autostart_win_only"))
            return
        try:
            items = read_autostart()
        except Exception as e:
            self.lbl_auto.config(text=f"{t('error_title')}: {e}")
            return
        active = sum(1 for i in items if i[1] == t("status_active"))
        missing = 0
        for name, status, command, source in items:
            target = target_status(command)
            hint = autostart_hint(name, command, target)
            tags = ()
            if target == t("status_missing"):
                missing += 1
                tags = ("missing",)
            elif hint.startswith("⚠"):
                tags = ("warn",)
            elif status != t("status_active"):
                tags = ("off",)
            self.tree_auto.add(name, (status, target, hint, command, source), tags=tags)
        summary = t("lbl_active_entries").format(active, len(items))
        if missing:
            summary += t("lbl_missing_files").format(missing)
        self.lbl_auto_summary.config(text=summary)
        self.lbl_auto.config(text=t("lbl_autostart_only_display"))

    def open_taskmgr(self):
        try:
            subprocess.Popen(["taskmgr"])
        except Exception as e:
            messagebox.showerror(t("error_title"), t("error_taskmgr_failed").format(e))

    def open_startup_settings(self):
        try:
            os.startfile("ms-settings:startupapps")
        except Exception as e:
            messagebox.showerror(t("error_title"), t("error_settings_failed").format(e))

    def build_folders_tab(self):
        bar = tk.Frame(self.tab_folders, bg=BG)
        bar.pack(fill=tk.X)
        self.btn_scan = ttk.Button(bar, text=t("btn_choose"), style="Primary.TButton",
                                   command=self.choose_and_scan)
        self.btn_scan.pack(side=tk.LEFT)
        self.btn_up = ttk.Button(bar, text=t("btn_up"), command=self.go_up, state=tk.DISABLED)
        self.btn_up.pack(side=tk.LEFT, padx=8)
        self.btn_cancel = ttk.Button(bar, text=t("btn_cancel"), command=self.on_cancel, state=tk.DISABLED)
        self.btn_cancel.pack(side=tk.LEFT)
        self.btn_export = ttk.Button(bar, text=t("btn_export"), command=self.export_results, state=tk.DISABLED)
        self.btn_export.pack(side=tk.RIGHT)
        info = make_card(self.tab_folders)
        info.pack(fill=tk.X, pady=12)
        inner = tk.Frame(info, bg=CARD)
        inner.pack(fill=tk.X, padx=18, pady=14)
        self.lbl_folder = tk.Label(inner, text=t("lbl_folder_not_chosen"), font=FONT_BOLD, bg=CARD, fg=TEXT,
                                   anchor="w")
        self.lbl_folder.pack(fill=tk.X)
        self.lbl_status = tk.Label(inner, text=t("lbl_folder_hint"),
                                   font=FONT_SMALL, bg=CARD, fg=MUTED, anchor="w", justify="left")
        self.lbl_status.pack(fill=tk.X, pady=(3, 8))
        auto_wrap(self.lbl_status, inner, 10)
        self.progress = ttk.Progressbar(inner, style="Accent.Horizontal.TProgressbar", mode="determinate")
        self.progress.pack(fill=tk.X)
        self.results_box = tk.Frame(self.tab_folders, bg=BG)
        self.results_box.pack(fill=tk.BOTH, expand=True)
        self.empty_label = make_label(
            self.results_box, t("lbl_no_analysis"), FONT, MUTED, justify="center")
        self.empty_label.pack(expand=True)
        self.res_nb = ttk.Notebook(self.results_box)
        tab_f = tk.Frame(self.res_nb, bg=BG, pady=10)
        tab_files = tk.Frame(self.res_nb, bg=BG, pady=10)
        tab_types = tk.Frame(self.res_nb, bg=BG, pady=10)
        self.res_nb.add(tab_f, text=t("tab_folder"))
        self.res_nb.add(tab_files, text=t("tab_files"))
        self.res_nb.add(tab_types, text=t("tab_types"))
        self.tree_folders = make_tree(
            tab_f, (("#0", t("col_folder"), 300, "w"), ("groesse", t("col_size"), 90, "e"), ("anteil", t("col_share"), 70, "e"),
                    ("neueste", t("col_newest"), 140, "center"), ("alter", t("col_age"), 85, "e"),
                    ("hinweis", t("col_hint"), 300, "w")),
            numeric=("groesse", "anteil", "neueste", "alter"), height=14)
        self.tree_files = make_tree(
            tab_files, (("#0", t("col_file"), 340, "w"), ("groesse", t("col_size"), 90, "e"),
                        ("datum", t("col_changed"), 140, "center"), ("alter", t("col_age"), 85, "e"),
                        ("hinweis", t("col_hint"), 300, "w")),
            numeric=("groesse", "datum", "alter"), height=14)
        self.tree_types = make_tree(
            tab_types, (("#0", t("col_type"), 250, "w"), ("groesse", t("col_size"), 110, "e"), ("anteil", t("col_share"), 80, "e")),
            numeric=("groesse", "anteil"), height=14)
        for tree in (self.tree_folders, self.tree_files):
            tree.tag_configure("sys", foreground="#b23b3b")
            tree.tag_configure("warn", foreground="#b26a00")
        self.attach_menu(self.tree_folders, folder_tree=True)
        self.attach_menu(self.tree_files, folder_tree=False)
        self.tree_folders.bind("<Double-1>", lambda e: self.open_row(self.tree_folders, e, True))
        self.tree_files.bind("<Double-1>", lambda e: self.open_row(self.tree_files, e, False))

    def show_empty(self, text):
        self.res_nb.pack_forget()
        self.empty_label.config(text=text)
        self.empty_label.pack(expand=True)

    def show_results_view(self):
        self.empty_label.pack_forget()
        self.res_nb.pack(fill=tk.BOTH, expand=True)

    def attach_menu(self, tree, folder_tree):
        menu = tk.Menu(self.root, tearoff=0)
        def selected():
            sel = tree.selection()
            return tree.paths.get(sel[0]) if sel else None
        def do_open():
            path = selected()
            if path and self.confirm_open(path):
                open_folder(path) if folder_tree else reveal_file(path)
        def do_copy():
            path = selected()
            if path:
                self.root.clipboard_clear()
                self.root.clipboard_append(path)
        def do_drill():
            path = selected()
            if path:
                self.start_scan(path)
        menu.add_command(label=t("open_in_explorer") if folder_tree else t("show_in_explorer"), command=do_open)
        menu.add_command(label=t("copy_path"), command=do_copy)
        if folder_tree:
            menu.add_separator()
            menu.add_command(label=t("analyze_here"), command=do_drill)
        def popup(event):
            iid = tree.identify_row(event.y)
            if iid and iid in tree.paths:
                tree.selection_set(iid)
                menu.tk_popup(event.x_root, event.y_root)
        tree.bind("<Button-3>", popup)
        tree.bind("<Button-2>", popup)

    def open_row(self, tree, event, folder_tree):
        iid = tree.identify_row(event.y)
        path = tree.paths.get(iid)
        if path and self.confirm_open(path):
            open_folder(path) if folder_tree else reveal_file(path)

    def show_safety_notice(self, startup=False):
        hidden = bool(self.settings.get("hide_startup_notice"))
        if startup and hidden:
            return
        dialog = SafetyDialog(self.root, hidden)
        self.root.wait_window(dialog)
        if dialog.hide_next != hidden:
            self.settings["hide_startup_notice"] = dialog.hide_next
            save_settings(self.settings)

    def confirm_open(self, path):
        info = classify_path(path)
        if not info:
            return True
        level, text = info
        return messagebox.askokcancel(
            t("safety_warning_title"),
            t("safety_warning_text").format(text),
            icon="warning", default="cancel" if level == LEVEL_SYSTEM else "ok")

    def open_storage_settings(self):
        try:
            os.startfile("ms-settings:storagesense")
        except Exception as e:
            messagebox.showerror(t("error_title"), t("error_settings_failed").format(e))

    @staticmethod
    def hint_for(path, locations):
        info = classify_path(path, locations)
        if not info:
            return "", ()
        return info[1], (("sys",) if info[0] == LEVEL_SYSTEM else ("warn",))

    def choose_and_scan(self):
        folder = filedialog.askdirectory(title=t("btn_choose"))
        if folder:
            self.start_scan(folder)

    def go_up(self):
        parent = os.path.dirname(self.scan_path)
        if parent and parent != self.scan_path:
            self.start_scan(parent)

    def start_scan(self, folder):
        self.show_tab(2)
        if self.busy:
            return
        folder = os.path.abspath(folder)
        if not os.path.isdir(folder):
            messagebox.showerror(t("error_title"), t("error_folder_not_found").format(folder))
            return
        self.scan_path = folder
        self.result = None
        for tree in (self.tree_folders, self.tree_files, self.tree_types):
            tree.clear()
        info = classify_path(folder)
        if info:
            self.lbl_folder.config(text=f"{folder}   {info[1]}",
                                   fg="#b23b3b" if info[0] == LEVEL_SYSTEM else "#b26a00")
        else:
            self.lbl_folder.config(text=folder, fg=TEXT)
        self.show_empty(t("lbl_analyzing"))
        self.cancel.clear()
        self.set_busy(True)
        threading.Thread(target=self.folder_worker, args=(folder,), daemon=True).start()

    def folder_worker(self, folder):
        try:
            self.post("result", scan_folder(folder, self.cancel, self.post))
        except Cancelled:
            self.post("status", t("lbl_analysis_aborted"))
            self.post("empty", t("lbl_analysis_aborted"))
        except Exception as e:
            self.post("status", f"{t('error_title')}: {e}")
            self.post("empty", t("lbl_analysis_failed"))
            self.post("error", traceback.format_exc())
        finally:
            self.post("idle")

    def on_cancel(self):
        self.cancel.set()
        self.lbl_status.config(text=t("lbl_canceling"))

    def on_close(self):
        self.cancel.set()
        self.root.destroy()

    def set_busy(self, busy):
        self.busy = busy
        self.btn_scan.config(state=tk.DISABLED if busy else tk.NORMAL)
        self.btn_cancel.config(state=tk.NORMAL if busy else tk.DISABLED)
        can_up = (not busy and bool(self.scan_path)
                  and os.path.dirname(self.scan_path) != self.scan_path)
        self.btn_up.config(state=tk.NORMAL if can_up else tk.DISABLED)
        self.btn_export.config(state=tk.NORMAL if (not busy and self.result) else tk.DISABLED)

    def post(self, *msg):
        self.q.put(msg)

    def poll_queue(self):
        try:
            for _ in range(200):
                self.handle(self.q.get_nowait())
        except queue.Empty:
            pass
        self.root.after(50, self.poll_queue)

    def handle(self, msg):
        kind = msg[0]
        if kind == "status":
            self.lbl_status.config(text=msg[1])
        elif kind == "progress":
            self.progress["maximum"] = max(msg[2], 1)
            self.progress["value"] = msg[1]
        elif kind == "result":
            self.show_results(msg[1])
        elif kind == "empty":
            self.show_empty(msg[1])
        elif kind == "error":
            messagebox.showerror(t("error_unexpected"), msg[1])
        elif kind == "idle":
            self.set_busy(False)

    def show_results(self, r):
        self.result = r
        total = r["total"] or 1
        for tree in (self.tree_folders, self.tree_files, self.tree_types):
            tree.clear()
        locations = protected_locations()
        for name, path, size, newest in r["folders"][:TOP_FOLDERS]:
            hint, tags = self.hint_for(path, locations)
            self.tree_folders.add(
                name, (fmt_size(size), fmt_pct(size / total * 100), fmt_date(newest), fmt_age(newest), hint),
                keys={"groesse": size, "anteil": size, "neueste": newest, "alter": age_seconds(newest)},
                path=path, tags=tags)
        if r["root_files"]:
            size, newest = r["root_files"], r["root_newest"]
            self.tree_folders.add(
                t("lbl_files_in_root"),
                (fmt_size(size), fmt_pct(size / total * 100), fmt_date(newest), fmt_age(newest), ""),
                keys={"groesse": size, "anteil": size, "neueste": newest, "alter": age_seconds(newest)})
        for size, path, mtime in r["files"]:
            hint, tags = self.hint_for(path, locations)
            self.tree_files.add(path, (fmt_size(size), fmt_date(mtime), fmt_age(mtime), hint),
                                keys={"groesse": size, "datum": mtime, "alter": age_seconds(mtime)},
                                path=path, tags=tags)
        for ext, size in r["ext"][:TOP_TYPES]:
            self.tree_types.add(ext, (fmt_size(size), fmt_pct(size / total * 100)),
                                keys={"groesse": size, "anteil": size})
        for tree in (self.tree_folders, self.tree_files, self.tree_types):
            tree.sort_by("groesse", True)
        self.show_results_view()
        self.progress["value"] = self.progress["maximum"]
        text = t("lbl_finished").format(fmt_size(r['total']), fmt_int(r['file_count']))
        if r["errors"]:
            text += t("lbl_not_readable").format(fmt_int(r['errors']))
        self.lbl_status.config(text=text)

    def export_results(self):
        if not self.result:
            return
        name = os.path.basename(self.result["folder"].rstrip("\\/")) or (t("unit_drive") if LANG == "en" else "Laufwerk")
        name = re.sub(r"[^\w.-]+", "_", name)
        stamp = time.strftime("%Y-%m-%d")
        path = filedialog.asksaveasfilename(
            defaultextension=".csv", initialfile=t("export_default_name").format(name, stamp),
            filetypes=[(t("export_csv"), "*.csv"), (t("export_txt"), "*.txt")])
        if not path:
            return
        try:
            if path.lower().endswith(".txt"):
                with open(path, "w", encoding="utf-8") as f:
                    f.write(build_text_report(self.result))
            else:
                with open(path, "w", newline="", encoding="utf-8-sig") as f:
                    csv.writer(f, delimiter=";").writerows(build_csv_rows(self.result))
        except OSError as e:
            messagebox.showerror(t("error_title"), t("error_export_failed").format(e))
            return
        messagebox.showinfo(t("error_exported"), t("error_exported_text").format(path))


if __name__ == "__main__":
    root = tk.Tk()
    AnalyseApp(root)
    root.mainloop()
