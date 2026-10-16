import os
import json

CONFIG_FILE = os.path.join(os.path.dirname(__file__), "team_config.json")

DEFAULT_CONFIG = {
    "team_members": ["Kaif", "Arhama", "Aliba"],
    "active_member": "Kaif"
}

def load_config():
    if not os.path.exists(CONFIG_FILE):
        save_config(DEFAULT_CONFIG)
        return DEFAULT_CONFIG.copy()
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            cfg = json.load(f)
            if not cfg.get("team_members"):
                cfg["team_members"] = ["Kaif", "Arhama", "Aliba"]
            if not cfg.get("active_member") or cfg["active_member"] not in cfg["team_members"]:
                cfg["active_member"] = "Kaif"
            return cfg
    except Exception:
        save_config(DEFAULT_CONFIG)
        return DEFAULT_CONFIG.copy()

def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2)

def add_member(name):
    clean_name = str(name).strip()
    if not clean_name:
        return False
    cfg = load_config()
    if clean_name not in cfg["team_members"]:
        cfg["team_members"].append(clean_name)
        save_config(cfg)
        return True
    return False

def remove_member(name):
    cfg = load_config()
    if name in cfg["team_members"]:
        cfg["team_members"].remove(name)
        if cfg["active_member"] == name:
            cfg["active_member"] = cfg["team_members"][0] if cfg["team_members"] else "Kaif"
        save_config(cfg)
        return True
    return False

def set_active_member(name):
    cfg = load_config()
    if name in cfg["team_members"]:
        cfg["active_member"] = name
        save_config(cfg)
        return True
    return False
