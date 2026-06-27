"""Storage for user-defined Action Plan templates, keyed by exclusion code.

Templates are stored as plain JSON next to the executable so the person can
edit the file directly if they prefer, in addition to the in-app editor.

File format:
{
  "Dvexclusion_001": "Continue routine follow up post-parathyroidectomy...",
  "VAC053": "Refer to vascular surgery for re-evaluation..."
}
"""
import json
import os
import sys


def _base_dir():
    """Directory the templates file should live in. For a frozen exe this is
    the exe's own folder (NOT sys._MEIPASS, which is a temp extraction dir
    that disappears after the program closes) so saved templates persist
    across runs. For a script run, it's the project root (one level up from
    core/)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


TEMPLATES_PATH = os.path.join(_base_dir(), "action_plan_templates.json")


def load_templates():
    """Return {code: plan_text} dict. Empty dict if the file doesn't exist
    or is unreadable (never raises)."""
    try:
        with open(TEMPLATES_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            return {str(k): str(v) for k, v in data.items()}
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        pass
    return {}


def save_templates(templates):
    """Write {code: plan_text} dict to disk as pretty JSON. Returns True/False."""
    try:
        with open(TEMPLATES_PATH, "w", encoding="utf-8") as f:
            json.dump(templates, f, ensure_ascii=False, indent=2, sort_keys=True)
        return True
    except OSError:
        return False


def set_template(code, plan_text):
    templates = load_templates()
    plan_text = (plan_text or "").strip()
    if plan_text:
        templates[code] = plan_text
    else:
        templates.pop(code, None)
    return save_templates(templates)
