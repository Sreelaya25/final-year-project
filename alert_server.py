from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse, RedirectResponse, JSONResponse
from pydantic import BaseModel
from typing import Optional
import random
import threading
import time
import json
from datetime import datetime
import os

app = FastAPI(title="AcciSense Emergency Dashboard")

STATE_FILE = "accisense_state.json"

def load_state():
    defaults = {
        "current_case": None,
        "hospital_index": 0,
        "selected_police": None,
        "case_status": "idle",
        "alert_time": None,
        "accept_time": None,
        "case_history": [],
        "hospital_stats": {h["name"]: {"accepted": 0, "declined": 0, "times": []} for h in HOSPITALS},
        "police_stats": {p: {"cases": 0} for p in POLICE_STATIONS},
        "total_cases": 0,
        "terminal_log": [],
    }
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
            for k, v in saved.items():
                defaults[k] = v
            for h in HOSPITALS:
                if h["name"] not in defaults["hospital_stats"]:
                    defaults["hospital_stats"][h["name"]] = {"accepted": 0, "declined": 0, "times": []}
            for p in POLICE_STATIONS:
                if p not in defaults["police_stats"]:
                    defaults["police_stats"][p] = {"cases": 0}
            print(f"[AcciSense] Loaded: {defaults['total_cases']} total cases")
        except Exception as e:
            print(f"[AcciSense] Could not load state ({e}), starting fresh.")
    else:
        print("[AcciSense] No existing state file, starting fresh.")
    return defaults

def save_state():
    try:
        with open(STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"[AcciSense] WARNING: Could not save state: {e}")

voice_thread = None
stop_voice_flag = False

def voice_loop():
    global stop_voice_flag
    try:
        import pythoncom, win32com.client
        pythoncom.CoInitialize()
        speaker = win32com.client.Dispatch("SAPI.SpVoice")
        while not stop_voice_flag:
            speaker.Speak("Emergency! Accident Detected! Emergency!")
            time.sleep(2)
        pythoncom.CoUninitialize()
    except:
        while not stop_voice_flag:
            print("🔊 VOICE ALERT: Emergency! Accident Detected!")
            time.sleep(2)

def start_voice_alert():
    global voice_thread, stop_voice_flag
    stop_voice_flag = False
    voice_thread = threading.Thread(target=voice_loop, daemon=True)
    voice_thread.start()

def stop_voice_alert():
    global stop_voice_flag
    stop_voice_flag = True

POLICE_STATIONS = [
    "Central Police Station","North Traffic Police HQ","South City Police Station",
    "East Zone Police Station","West District Control Room","Highway Patrol Unit Alpha",
    "Metro Traffic Command","Airport Zone Police Station","Industrial Area Police Post","City Control Room Central"
]

HOSPITALS = [
    {"name": "Apollo Emergency Center", "distance": "1.2 km"},
    {"name": "City Care Hospital", "distance": "2.4 km"},
    {"name": "Metro Trauma Hospital", "distance": "3.1 km"},
    {"name": "Green Cross Medical", "distance": "4.8 km"},
    {"name": "National Emergency Hospital", "distance": "6.2 km"},
]

VEHICLE_TYPES = {
    "TN01":{"type":"Sedan","icon":"🚗"},"TN02":{"type":"SUV","icon":"🚙"},"TN03":{"type":"Truck","icon":"🚛"},
    "MH":{"type":"Motorcycle","icon":"🏍️"},"KA":{"type":"Auto-Rickshaw","icon":"🛺"},"AP":{"type":"Bus","icon":"🚌"},
    "TS":{"type":"Van","icon":"🚐"},"DL":{"type":"Hatchback","icon":"🚘"},"GJ":{"type":"Pickup Truck","icon":"🛻"},
    "RJ":{"type":"Minibus","icon":"🚎"},"HR":{"type":"Bicycle","icon":"🚲"},"UP":{"type":"Ambulance","icon":"🚑"},
    "WB":{"type":"Taxi","icon":"🚕"},"MP":{"type":"Police Vehicle","icon":"🚓"},"PB":{"type":"Fire Truck","icon":"🚒"},
}

BED_STATUS = [
    {"ward":"Emergency ICU","total":20,"occupied":14,"available":6,"status":"warning"},
    {"ward":"General Ward A","total":50,"occupied":32,"available":18,"status":"good"},
    {"ward":"General Ward B","total":50,"occupied":45,"available":5,"status":"critical"},
    {"ward":"Trauma Unit","total":15,"occupied":8,"available":7,"status":"good"},
    {"ward":"Pediatric Ward","total":25,"occupied":10,"available":15,"status":"good"},
    {"ward":"Orthopedic Ward","total":30,"occupied":22,"available":8,"status":"warning"},
    {"ward":"Cardiac ICU","total":12,"occupied":11,"available":1,"status":"critical"},
    {"ward":"Neurology Ward","total":20,"occupied":13,"available":7,"status":"good"},
]

FLEET_STATUS = [
    {"id":"AMB-001","type":"Advanced Life Support","status":"available","driver":"Rajan Kumar","location":"Base Station","last_trip":"10:15 AM"},
    {"id":"AMB-002","type":"Basic Life Support","status":"dispatched","driver":"Suresh Patel","location":"En Route - NH44","last_trip":"11:30 AM"},
    {"id":"AMB-003","type":"Patient Transport","status":"available","driver":"Anil Sharma","location":"Base Station","last_trip":"09:45 AM"},
    {"id":"AMB-004","type":"Advanced Life Support","status":"maintenance","driver":"N/A","location":"Service Bay","last_trip":"08:00 AM"},
    {"id":"AMB-005","type":"Neonatal Transport","status":"available","driver":"Priya Nair","location":"Base Station","last_trip":"07:30 AM"},
    {"id":"AMB-006","type":"Basic Life Support","status":"dispatched","driver":"Vikram Singh","location":"City Hospital","last_trip":"11:45 AM"},
    {"id":"POL-001","type":"Police Patrol","status":"available","driver":"Officer Mehta","location":"Zone 3","last_trip":"10:00 AM"},
    {"id":"POL-002","type":"Traffic Control","status":"dispatched","driver":"Officer Das","location":"Accident Site","last_trip":"11:30 AM"},
]

state = load_state()

def log(msg):
    timestamp = datetime.now().strftime("%H:%M:%S")
    entry = f"[{timestamp}] {msg}"
    state["terminal_log"].append(entry)
    print(entry)
    if len(state["terminal_log"]) > 200:
        state["terminal_log"] = state["terminal_log"][-200:]
    save_state()

class AccidentAlert(BaseModel):
    vehicle_1: str
    vehicle_2: str
    latitude: Optional[float] = None
    longitude: Optional[float] = None

def get_vehicle_info(plate: str):
    for prefix, info in VEHICLE_TYPES.items():
        if plate.upper().startswith(prefix):
            return info
    return {"type": "Unknown Vehicle", "icon": "🚗"}

def current_hospital():
    idx = state["hospital_index"]
    if idx >= len(HOSPITALS):
        return None
    return HOSPITALS[idx]

@app.post("/alert")
def receive_alert(alert: AccidentAlert):
    state["current_case"] = alert.dict()
    state["hospital_index"] = 0
    state["selected_police"] = random.choice(POLICE_STATIONS)
    state["case_status"] = "pending"
    state["alert_time"] = datetime.now().isoformat()
    state["accept_time"] = None
    state["total_cases"] += 1
    police = state["selected_police"]
    state["police_stats"][police]["cases"] += 1
    v1info = get_vehicle_info(alert.vehicle_1)
    v2info = get_vehicle_info(alert.vehicle_2)
    save_state()
    log(f"🚨 ACCIDENT — V1: {alert.vehicle_1} ({v1info['type']}) | V2: {alert.vehicle_2} ({v2info['type']})")
    log(f"📍 GPS: {alert.latitude}, {alert.longitude}")
    log(f"🚓 Police: {police}")
    log(f"🏥 Hospital: {HOSPITALS[0]['name']}")
    start_voice_alert()
    return {"message": "Alert received", "redirect": "/police"}

@app.post("/accept")
def accept_case():
    state["case_status"] = "ambulance_dispatched"
    state["accept_time"] = datetime.now().isoformat()
    stop_voice_alert()
    hosp = current_hospital()
    if hosp:
        hname = hosp["name"]
        alert_time = datetime.fromisoformat(state["alert_time"])
        accept_time = datetime.fromisoformat(state["accept_time"])
        response_secs = int((accept_time - alert_time).total_seconds())
        state["hospital_stats"][hname]["accepted"] += 1
        state["hospital_stats"][hname]["times"].append(response_secs)
        log(f"✅ {hname} ACCEPTED — Response: {response_secs}s")
        c = state["current_case"]
        v1info = get_vehicle_info(c["vehicle_1"])
        v2info = get_vehicle_info(c["vehicle_2"])
        state["case_history"].append({
            "date": datetime.now().strftime("%Y-%m-%d"),
            "time": datetime.now().strftime("%H:%M:%S"),
            "v1": c["vehicle_1"],"v2": c["vehicle_2"],
            "v1type": v1info["type"],"v2type": v2info["type"],
            "v1icon": v1info["icon"],"v2icon": v2info["icon"],
            "hospital": hname,"police": state["selected_police"],
            "status": "Dispatched","response_time": response_secs
        })
        save_state()
    return RedirectResponse(url="/hospital", status_code=303)

@app.post("/decline")
def decline_case():
    hosp = current_hospital()
    if hosp:
        state["hospital_stats"][hosp["name"]]["declined"] += 1
        log(f"❌ {hosp['name']} DECLINED")
    state["hospital_index"] += 1
    next_hosp = current_hospital()
    if next_hosp is None:
        state["case_status"] = "no_hospital"
        stop_voice_alert()
        log("⛔ No hospitals available")
    else:
        log(f"🏥 Contacting: {next_hosp['name']}")
    save_state()
    return RedirectResponse(url="/hospital", status_code=303)

@app.post("/reset")
def reset_all():
    global state
    state.update({
        "current_case":None,"hospital_index":0,"selected_police":None,
        "case_status":"idle","alert_time":None,"accept_time":None,"case_history":[],
        "hospital_stats":{h["name"]:{"accepted":0,"declined":0,"times":[]} for h in HOSPITALS},
        "police_stats":{p:{"cases":0} for p in POLICE_STATIONS},
        "total_cases":0,"terminal_log":[]
    })
    save_state()
    log("🔄 State reset")
    return {"message": "State reset successfully"}

@app.get("/ack")
def ack():
    return JSONResponse({"accepted": state["case_status"] == "ambulance_dispatched"})

@app.get("/api/state")
def get_state():
    return JSONResponse(state)

@app.get("/api/logs")
def get_logs():
    return JSONResponse({"logs": state["terminal_log"]})

@app.get("/api/hospital-stats")
def hospital_stats():
    stats = []
    for h in HOSPITALS:
        name = h["name"]
        s = state["hospital_stats"][name]
        total = s["accepted"] + s["declined"]
        avg_time = int(sum(s["times"]) / len(s["times"])) if s["times"] else 0
        decline_pct = round((s["declined"] / total * 100), 1) if total > 0 else 0
        stats.append({"name":name,"accepted":s["accepted"],"declined":s["declined"],"total":total,"avg_time":avg_time,"decline_pct":decline_pct})
    return JSONResponse(stats)

@app.get("/api/bed-status")
def get_bed_status():
    return JSONResponse(BED_STATUS)

@app.get("/api/fleet-status")
def get_fleet_status():
    return JSONResponse(FLEET_STATUS)

@app.get("/", response_class=HTMLResponse)
def login_page():
    return HTMLResponse(LOGIN_HTML)

@app.post("/login")
async def login(request: Request):
    form = await request.form()
    role = form.get("role")
    if role == "police":
        return RedirectResponse(url="/police", status_code=303)
    elif role == "hospital":
        return RedirectResponse(url="/hospital", status_code=303)
    return RedirectResponse(url="/", status_code=303)

@app.get("/police", response_class=HTMLResponse)
def police_dashboard():
    return HTMLResponse(POLICE_HTML)

@app.get("/hospital", response_class=HTMLResponse)
def hospital_dashboard():
    return HTMLResponse(HOSPITAL_HTML)

# ─────────────────────────────────────────────────────────────────
# SHARED NOTIFICATION + BELL JS (injected into both dashboards)
# ─────────────────────────────────────────────────────────────────
BELL_CSS = """
/* ── BELL WRAP ── */
.bell-wrap{position:relative;cursor:pointer;display:flex;align-items:center;}
.bell-btn{width:36px;height:36px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:19px;cursor:pointer;transition:background 0.18s;background:none;border:none;padding:0;line-height:1;}
.bell-btn:hover{background:#f1f5f9;}
.bell-badge{position:absolute;top:-4px;right:-4px;min-width:16px;height:16px;background:#ef4444;border-radius:8px;border:2px solid white;display:flex;align-items:center;justify-content:center;font-size:8px;color:white;font-weight:900;padding:0 3px;}
.bell-badge.hidden{display:none;}
/* ── NOTIFICATION DROPDOWN ── */
.notif-panel{display:none;position:fixed;top:66px;right:16px;width:330px;background:#fff;border:1px solid #e2e8f0;border-radius:16px;box-shadow:0 20px 60px rgba(0,0,0,0.15);z-index:9999;overflow:hidden;flex-direction:column;}
.notif-panel.open{display:flex;animation:ndrop .18s ease;}
@keyframes ndrop{from{opacity:0;transform:translateY(-10px);}to{opacity:1;transform:translateY(0);}}
.notif-hdr{display:flex;align-items:center;justify-content:space-between;padding:14px 16px 11px;border-bottom:1px solid #f1f5f9;}
.notif-hdr-title{font-size:13px;font-weight:900;color:#0f172a;display:flex;align-items:center;gap:6px;}
.notif-mark{font-size:11px;font-weight:700;color:var(--accent,#3b82f6);cursor:pointer;background:none;border:none;font-family:'Nunito',sans-serif;padding:0;}
.notif-mark:hover{text-decoration:underline;}
.notif-scroll{max-height:310px;overflow-y:auto;}
.notif-item{display:flex;gap:11px;padding:11px 14px;border-bottom:1px solid #f8fafc;cursor:pointer;transition:background .15s;position:relative;}
.notif-item:hover{background:#f8fafc;}
.notif-item.unread{background:var(--notif-unread,#eff6ff);}
.notif-item.unread::before{content:'';position:absolute;left:0;top:0;bottom:0;width:3px;background:var(--accent,#3b82f6);border-radius:0 2px 2px 0;}
.notif-ic{width:34px;height:34px;border-radius:9px;display:flex;align-items:center;justify-content:center;font-size:15px;flex-shrink:0;margin-top:1px;}
.notif-body{}
.notif-ttl{font-size:12px;font-weight:800;color:#0f172a;margin-bottom:2px;}
.notif-desc{font-size:11px;color:#64748b;font-weight:600;line-height:1.4;}
.notif-ts{font-size:10px;color:#94a3b8;margin-top:3px;font-weight:700;}
.notif-empty{padding:36px;text-align:center;}
.notif-empty-ico{font-size:30px;opacity:.2;margin-bottom:8px;}
.notif-empty-txt{font-size:12px;color:#94a3b8;font-weight:700;}
.notif-ftr{padding:10px 16px;border-top:1px solid #f1f5f9;text-align:center;}
.notif-ftr a{font-size:11px;font-weight:800;color:var(--accent,#3b82f6);text-decoration:none;}
.notif-ftr a:hover{text-decoration:underline;}
/* ── GOOGLE MODAL ── */
.g-overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:99999;align-items:center;justify-content:center;}
.g-overlay.open{display:flex;}
.g-modal{background:#fff;border-radius:18px;padding:36px 30px 28px;width:350px;box-shadow:0 24px 70px rgba(0,0,0,.28);text-align:center;}
.g-modal-ico{font-size:36px;margin-bottom:12px;}
.g-modal h3{font-size:18px;font-weight:900;color:#111;margin-bottom:6px;}
.g-modal p{font-size:12px;color:#888;font-weight:600;margin-bottom:20px;line-height:1.6;}
.g-info-box{background:#f8faff;border:1px solid #e0e7ff;border-radius:10px;padding:12px 14px;text-align:left;margin-bottom:18px;font-size:11px;color:#4338ca;font-weight:700;line-height:1.5;}
.g-signin-btn{width:100%;padding:13px;background:#4285F4;border:none;border-radius:10px;color:#fff;font-family:'Nunito',sans-serif;font-size:13px;font-weight:800;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:9px;margin-bottom:10px;transition:all .2s;}
.g-signin-btn:hover{background:#3367d6;transform:translateY(-1px);box-shadow:0 6px 18px rgba(66,133,244,.4);}
.g-signin-ico{width:20px;height:20px;border-radius:50%;background:#fff;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:900;color:#4285F4;flex-shrink:0;}
.g-cancel-btn{width:100%;padding:10px;background:none;border:1.5px solid #e5e7eb;border-radius:10px;font-family:'Nunito',sans-serif;font-size:12px;font-weight:700;color:#888;cursor:pointer;}
.g-cancel-btn:hover{background:#f9f9f9;}
"""

BELL_JS = """
// ── NOTIFICATION BELL ──
var _notifOpen=false;
var _unread=0;
var _notifItems=[];

function _initBell(accentColor,unreadColor,initialItems){
  _notifItems=initialItems||[];
  _unread=_notifItems.filter(function(x){return x.unread;}).length;
  _renderNotifs();
  _updateBadge();
  document.addEventListener('click',function(e){
    var panel=document.getElementById('notifPanel');
    var wrap=document.getElementById('bellWrap');
    if(_notifOpen&&panel&&wrap&&!panel.contains(e.target)&&!wrap.contains(e.target)){
      _closeNotif();
    }
  });
}

function _toggleNotif(e){
  e.stopPropagation();
  _notifOpen=!_notifOpen;
  var panel=document.getElementById('notifPanel');
  if(panel) panel.classList.toggle('open',_notifOpen);
}

function _closeNotif(){
  _notifOpen=false;
  var panel=document.getElementById('notifPanel');
  if(panel) panel.classList.remove('open');
}

function _markAllRead(){
  _notifItems.forEach(function(x){x.unread=false;});
  _unread=0;
  _updateBadge();
  _renderNotifs();
}

function _updateBadge(){
  var badge=document.getElementById('bellBadge');
  if(!badge) return;
  if(_unread>0){badge.textContent=_unread;badge.classList.remove('hidden');}
  else{badge.classList.add('hidden');}
}

function _renderNotifs(){
  var list=document.getElementById('notifList');
  if(!list) return;
  if(_notifItems.length===0){
    list.innerHTML='<div class="notif-empty"><div class="notif-empty-ico">🔕</div><div class="notif-empty-txt">No notifications</div></div>';
    return;
  }
  list.innerHTML=_notifItems.map(function(n,i){
    return '<div class="notif-item'+(n.unread?' unread':'') +'" onclick="_readNotif('+i+')">'
      +'<div class="notif-ic" style="background:'+n.bg+'">'+n.icon+'</div>'
      +'<div class="notif-body">'
      +'<div class="notif-ttl">'+n.title+'</div>'
      +'<div class="notif-desc">'+n.desc+'</div>'
      +'<div class="notif-ts">'+n.time+'</div>'
      +'</div></div>';
  }).join('');
}

function _readNotif(i){
  if(_notifItems[i].unread){_notifItems[i].unread=false;_unread=Math.max(0,_unread-1);_updateBadge();_renderNotifs();}
}

function _pushNotif(icon,bg,title,desc){
  _notifItems.unshift({icon:icon,bg:bg,title:title,desc:desc,time:'Just now',unread:true});
  _unread++;
  _updateBadge();
  _renderNotifs();
}

// ── GOOGLE MODAL ──
function openGoogleModal(){document.getElementById('gOverlay').classList.add('open');}
function closeGoogleModal(){document.getElementById('gOverlay').classList.remove('open');}
function doGoogleLogin(role){closeGoogleModal();window.location.href=role==='police'?'/police':'/hospital';}
"""

# ─────────────────────────────────────────────
# LOGIN PAGE
# ─────────────────────────────────────────────
LOGIN_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>AcciSense — Emergency Response System</title>
<link href="https://fonts.googleapis.com/css2?family=Nunito:wght@400;500;600;700;800;900&display=swap" rel="stylesheet">
<style>
*{margin:0;padding:0;box-sizing:border-box;}
body{font-family:'Nunito',sans-serif;min-height:100vh;background:#7c3aed;display:flex;align-items:center;justify-content:center;padding:20px;position:relative;overflow:hidden;}
.blob{position:fixed;border-radius:50%;background:rgba(255,255,255,0.1);pointer-events:none;}
.blob1{width:320px;height:320px;top:-80px;left:-80px;}.blob2{width:200px;height:200px;bottom:60px;left:40px;}.blob3{width:160px;height:160px;bottom:-40px;right:220px;}.blob4{width:260px;height:260px;top:80px;right:-60px;opacity:0.06;}
.plus{position:fixed;color:rgba(255,255,255,0.22);font-size:40px;font-weight:900;pointer-events:none;line-height:1;}
.plus1{top:55px;left:55px;}.plus2{bottom:75px;right:75px;}.plus3{top:50%;left:18px;}
.card-wrap{display:flex;width:100%;max-width:860px;min-height:520px;border-radius:24px;overflow:hidden;box-shadow:0 30px 80px rgba(0,0,0,0.4);position:relative;z-index:1;}
.left{flex:1;background:#fff;padding:44px 40px;display:flex;flex-direction:column;position:relative;}
.left-logo{display:flex;align-items:center;justify-content:center;gap:10px;margin-bottom:22px;}
.logo-icon{width:44px;height:44px;border-radius:12px;display:flex;align-items:center;justify-content:center;font-size:22px;color:white;}
.logo-icon.hosp{background:linear-gradient(135deg,#06b6d4,#2563eb);}.logo-icon.police{background:linear-gradient(135deg,#7c3aed,#dc2626);}
.welcome-title{text-align:center;font-size:22px;font-weight:900;color:#111;margin-bottom:24px;}
.lang-btn{position:absolute;top:18px;right:18px;background:#f3f4f6;border:1px solid #e5e7eb;border-radius:20px;padding:5px 12px;font-size:12px;font-weight:700;color:#555;cursor:pointer;}
.tabs{display:flex;background:#f3f0fa;border-radius:12px;padding:4px;margin-bottom:24px;}
.tab{flex:1;padding:10px;text-align:center;border-radius:9px;cursor:pointer;font-size:14px;font-weight:800;color:#9b8ec4;transition:all .25s;display:flex;align-items:center;justify-content:center;gap:6px;}
.tab.active{background:#fff;color:#5b21b6;box-shadow:0 2px 8px rgba(0,0,0,.1);}
.field{margin-bottom:16px;}
.field-label{font-size:11px;font-weight:800;color:#7c3aed;letter-spacing:.5px;margin-bottom:6px;text-transform:uppercase;}
.inp-wrap{position:relative;display:flex;align-items:center;}
.inp-icon{position:absolute;left:14px;color:#bbb;font-size:15px;}
.inp-wrap input{width:100%;padding:12px 14px 12px 40px;border:1.5px solid #e8e0f5;border-radius:10px;font-family:'Nunito',sans-serif;font-size:14px;color:#333;outline:none;transition:border .2s;background:#faf8ff;}
.inp-wrap input:focus{border-color:#7c3aed;background:#fff;}
.eye-btn{position:absolute;right:14px;cursor:pointer;color:#bbb;font-size:14px;background:none;border:none;padding:0;}
.extras{display:flex;align-items:center;justify-content:space-between;margin-bottom:20px;font-size:13px;}
.remember{display:flex;align-items:center;gap:7px;color:#555;font-weight:700;cursor:pointer;}
.remember input[type=checkbox]{accent-color:#7c3aed;width:15px;height:15px;}
.forgot{color:#7c3aed;font-weight:800;text-decoration:none;}
.forgot:hover{text-decoration:underline;}
.btn-login{width:100%;padding:14px;background:linear-gradient(135deg,#7c3aed,#9333ea);border:none;border-radius:12px;color:#fff;font-family:'Nunito',sans-serif;font-size:15px;font-weight:900;letter-spacing:1.5px;cursor:pointer;transition:all .25s;box-shadow:0 6px 20px rgba(124,58,237,.4);margin-bottom:14px;}
.btn-login:hover{transform:translateY(-2px);box-shadow:0 10px 28px rgba(124,58,237,.5);}
.or-line{display:flex;align-items:center;gap:10px;margin-bottom:14px;color:#bbb;font-size:13px;}
.or-line::before,.or-line::after{content:'';flex:1;height:1px;background:#eee;}
.btn-google{width:100%;padding:12px;background:#fff;border:1.5px solid #e0dbe9;border-radius:12px;color:#333;font-family:'Nunito',sans-serif;font-size:14px;font-weight:700;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:10px;transition:all .2s;margin-bottom:12px;}
.btn-google:hover{background:#f9f7ff;border-color:#c0b4e0;transform:translateY(-1px);}
.g-ico-wrap{width:22px;height:22px;border-radius:50%;background:linear-gradient(135deg,#4285F4 25%,#34A853 25%,#34A853 50%,#FBBC05 50%,#FBBC05 75%,#EA4335 75%);display:flex;align-items:center;justify-content:center;font-size:11px;color:white;font-weight:900;flex-shrink:0;}
.secure-gov{text-align:center;font-size:12px;color:#bbb;font-weight:700;}
.right{width:340px;background:linear-gradient(160deg,#5b21b6 0%,#7c3aed 45%,#8b5cf6 100%);padding:48px 32px;display:flex;flex-direction:column;align-items:center;justify-content:center;position:relative;overflow:hidden;}
.rb{position:absolute;border-radius:50%;background:rgba(255,255,255,.08);}
.rb1{width:180px;height:180px;top:-40px;right:-40px;}.rb2{width:120px;height:120px;bottom:60px;left:-30px;}.rb3{width:80px;height:80px;bottom:-20px;right:40px;}
.siren-box{width:90px;height:90px;background:rgba(255,255,255,.12);border-radius:22px;display:flex;align-items:center;justify-content:center;font-size:46px;margin-bottom:26px;border:1px solid rgba(255,255,255,.2);box-shadow:0 8px 32px rgba(0,0,0,.2);}
.right h2{color:#fff;font-size:22px;font-weight:900;text-align:center;margin-bottom:12px;}
.right p{color:rgba(255,255,255,.72);font-size:13px;text-align:center;line-height:1.75;margin-bottom:34px;}
.role-chips{display:flex;gap:12px;justify-content:center;}
.role-chip{background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.2);border-radius:14px;padding:14px 20px;text-align:center;color:#fff;cursor:pointer;transition:all .2s;min-width:90px;}
.role-chip:hover{background:rgba(255,255,255,.22);transform:translateY(-2px);}
.rc-icon{font-size:26px;margin-bottom:6px;}.rc-label{font-size:12px;font-weight:900;letter-spacing:.5px;}
</style>
<style>""" + BELL_CSS + """</style>
</head>
<body>
<div class="blob blob1"></div><div class="blob blob2"></div><div class="blob blob3"></div><div class="blob blob4"></div>
<div class="plus plus1">+</div><div class="plus plus2">+</div><div class="plus plus3">+</div>

<!-- GOOGLE MODAL -->
<div class="g-overlay" id="gOverlay" onclick="if(event.target===this)closeGoogleModal()">
  <div class="g-modal">
    <div class="g-modal-ico">🔐</div>
    <h3>Sign in with Google</h3>
    <p>AcciSense Government Portal<br>Secure access for authorized personnel only.</p>
    <div class="g-info-box">⚠️ This portal is restricted to pre-registered government accounts. Unauthorized access is prohibited and monitored.</div>
    <button class="g-signin-btn" onclick="doGoogleLogin(document.getElementById('roleInput').value)">
      <div class="g-signin-ico">G</div> Continue with Google Account
    </button>
    <button class="g-cancel-btn" onclick="closeGoogleModal()">Cancel</button>
  </div>
</div>

<div class="card-wrap">
  <div class="left">
    <div class="lang-btn">🇬🇧 GB English ▾</div>
    <div class="left-logo">
      <div class="logo-icon hosp">➕</div>
      <div class="logo-icon police">🛡️</div>
    </div>
    <div class="welcome-title">Welcome Back</div>
    <div class="tabs">
      <div class="tab active" id="tab-hosp" onclick="selRole('hospital')">➕ Hospital</div>
      <div class="tab" id="tab-police" onclick="selRole('police')">🛡️ Police</div>
    </div>
    <form method="post" action="/login">
      <input type="hidden" name="role" id="roleInput" value="hospital">
      <div class="field">
        <div class="field-label">Email Address</div>
        <div class="inp-wrap"><span class="inp-icon">👤</span><input type="text" placeholder="Enter your email" value="admin@accisense.gov"></div>
      </div>
      <div class="field">
        <div class="field-label">Password</div>
        <div class="inp-wrap"><span class="inp-icon">🔒</span><input type="password" id="pwdInput" placeholder="Enter your password" value="admin123"><button type="button" class="eye-btn" onclick="togglePwd()">👁</button></div>
      </div>
      <div class="extras">
        <label class="remember"><input type="checkbox" checked> Remember me</label>
        <a href="#" class="forgot">Forgot Password?</a>
      </div>
      <button class="btn-login" type="submit">LOGIN</button>
      <div class="or-line">or Login with OTP</div>
      <button class="btn-google" type="button" onclick="openGoogleModal()">
        <div class="g-ico-wrap">G</div> Continue with Google
      </button>
      <div class="secure-gov">🔒 Secure Government Access</div>
    </form>
  </div>
  <div class="right">
    <div class="rb rb1"></div><div class="rb rb2"></div><div class="rb rb3"></div>
    <div class="siren-box">🚨</div>
    <h2>Welcome Back!</h2>
    <p>AcciSense — Emergency Response<br>Intelligence Platform.<br>Connecting hospitals & police<br>for faster emergency response.</p>
    <div class="role-chips">
      <div class="role-chip" onclick="selRole('hospital')"><div class="rc-icon">➕</div><div class="rc-label">Hospital</div></div>
      <div class="role-chip" onclick="selRole('police')"><div class="rc-icon">🛡️</div><div class="rc-label">Police</div></div>
    </div>
  </div>
</div>
<script>
""" + BELL_JS + """
function selRole(r){document.getElementById('roleInput').value=r;document.getElementById('tab-hosp').classList.toggle('active',r==='hospital');document.getElementById('tab-police').classList.toggle('active',r==='police');}
function togglePwd(){var i=document.getElementById('pwdInput');i.type=i.type==='password'?'text':'password';}
</script>
</body>
</html>"""

# ─────────────────────────────────────────────
# POLICE DASHBOARD
# ─────────────────────────────────────────────
POLICE_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>AcciSense — Police Dashboard</title>
<link href="https://fonts.googleapis.com/css2?family=Nunito:wght@400;500;600;700;800;900&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<style>
*{margin:0;padding:0;box-sizing:border-box;}
:root{--bg:#f0f4f8;--sidebar:#1c2b4a;--surface:#fff;--border:#e2e8f0;--accent:#3b82f6;--warn:#f59e0b;--danger:#ef4444;--text:#0f172a;--muted:#64748b;--sw:195px;--notif-unread:#eff6ff;}
body{background:var(--bg);font-family:'Nunito',sans-serif;color:var(--text);display:flex;min-height:100vh;}
.sidebar{width:var(--sw);background:var(--sidebar);display:flex;flex-direction:column;position:fixed;top:0;left:0;height:100vh;z-index:100;overflow-y:auto;}
.sb-logo{padding:20px 16px 16px;border-bottom:1px solid rgba(255,255,255,.07);}
.sb-logo-row{display:flex;align-items:center;gap:10px;}
.sb-badge{width:42px;height:42px;background:linear-gradient(135deg,#2563eb,#1d4ed8);border-radius:12px;display:flex;align-items:center;justify-content:center;font-size:20px;flex-shrink:0;box-shadow:0 4px 12px rgba(37,99,235,.4);}
.sb-name{font-size:14px;font-weight:900;color:white;}.sb-sub{font-size:9px;color:#475569;letter-spacing:1.5px;margin-top:1px;text-transform:uppercase;}
.sb-section{padding:16px 16px 5px;font-size:9px;font-weight:800;letter-spacing:2px;color:#2d4a6d;text-transform:uppercase;}
.nav{display:flex;align-items:center;gap:9px;padding:9px 14px 9px 16px;cursor:pointer;transition:all .2s;font-size:12px;font-weight:700;color:#4a7a9b;margin:1px 0;border-right:3px solid transparent;}
.nav:hover{background:rgba(255,255,255,.05);color:#7fb3d3;}
.nav.active{background:rgba(59,130,246,.15);color:#60a5fa;border-right-color:#3b82f6;}
.nav-ic{font-size:13px;width:17px;text-align:center;}
.nav-badge{margin-left:auto;background:#ef4444;color:white;border-radius:20px;padding:1px 6px;font-size:9px;font-weight:800;}
.sb-bottom{margin-top:auto;padding:14px;border-top:1px solid rgba(255,255,255,.06);}
.sb-user{display:flex;align-items:center;gap:9px;}
.av{width:34px;height:34px;background:linear-gradient(135deg,#3b82f6,#1d4ed8);border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:13px;font-weight:900;color:white;}
.un{font-size:11px;font-weight:800;color:#e2e8f0;}.ur{font-size:9px;color:#4a6580;letter-spacing:1px;text-transform:uppercase;}
.topbar{position:fixed;top:0;left:var(--sw);right:0;height:58px;background:white;border-bottom:1px solid var(--border);display:flex;align-items:center;padding:0 24px;z-index:50;box-shadow:0 1px 3px rgba(0,0,0,.05);}
.tb-title{font-size:14px;font-weight:900;color:var(--text);}
.tb-right{margin-left:auto;display:flex;align-items:center;gap:12px;}
.live-pill{display:flex;align-items:center;gap:6px;background:#eff6ff;border:1px solid #bfdbfe;border-radius:20px;padding:5px 12px;font-size:11px;font-weight:800;color:#2563eb;letter-spacing:1px;}
.lp-dot{width:6px;height:6px;border-radius:50%;background:#3b82f6;animation:pulse 1.2s infinite;}
.tb-time{font-size:12px;color:var(--muted);font-weight:700;}
.tb-logout{background:#f8fafc;border:1px solid var(--border);color:var(--muted);padding:6px 12px;border-radius:8px;cursor:pointer;font-size:12px;font-weight:700;text-decoration:none;}
.main{margin-left:var(--sw);margin-top:58px;flex:1;padding:24px;overflow-y:auto;}
.page-title{font-size:20px;font-weight:900;color:var(--text);margin-bottom:3px;}
.page-sub{font-size:13px;color:var(--muted);margin-bottom:20px;font-weight:600;}
.persist-info{display:flex;align-items:center;justify-content:space-between;gap:10px;background:#eff6ff;border:1px solid #bfdbfe;border-radius:10px;padding:10px 16px;margin-bottom:18px;font-size:12px;color:#1e40af;font-weight:700;}
.persist-left{display:flex;align-items:center;gap:8px;}
.persist-dot{width:8px;height:8px;border-radius:50%;background:#3b82f6;flex-shrink:0;}
.stats-row{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:20px;}
.stat-card{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:18px;transition:box-shadow .2s;}
.stat-card:hover{box-shadow:0 8px 24px rgba(0,0,0,.08);}
.stat-top{display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:10px;}
.stat-label{font-size:11px;font-weight:700;color:var(--muted);}
.stat-icon-box{width:36px;height:36px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:16px;}
.stat-value{font-size:28px;font-weight:900;color:var(--text);line-height:1;}
.stat-sub{font-size:11px;color:var(--muted);margin-top:4px;font-weight:600;}
.disp-card-wrap{border-radius:16px;overflow:hidden;margin-bottom:20px;border:1.5px solid var(--border);background:var(--surface);}
.disp-card-wrap.emergency{border-color:#fca5a5;}.disp-card-wrap.resolved{border-color:#93c5fd;}
.disp-header{display:flex;align-items:center;gap:12px;padding:16px 20px;border-bottom:1px solid var(--border);}
.disp-pill-resolved{background:#eff6ff;border:1px solid #93c5fd;color:#1d4ed8;border-radius:8px;padding:5px 12px;font-size:11px;font-weight:800;display:flex;align-items:center;gap:5px;}
.disp-pill-emergency{background:#fef2f2;border:1px solid #fca5a5;color:#dc2626;border-radius:8px;padding:5px 12px;font-size:11px;font-weight:800;animation:flash 1.2s infinite;display:flex;align-items:center;gap:5px;}
.disp-title{font-size:14px;font-weight:800;color:var(--text);}
.disp-status-badge{margin-left:auto;padding:6px 16px;border-radius:20px;font-size:11px;font-weight:800;}
.dsb-blue{background:#1d4ed8;color:white;}.dsb-red{background:#dc2626;color:white;}.dsb-green{background:#16a34a;color:white;}
.wave-hero{position:relative;padding:24px 28px;min-height:120px;display:flex;align-items:center;gap:20px;overflow:hidden;}
.wave-hero.blue-wave{background:linear-gradient(135deg,#dbeafe,#bfdbfe,#93c5fd);}
.wave-hero.red-wave{background:linear-gradient(135deg,#fef2f2,#fee2e2,#fecaca);}
.wave-hero::after{content:'';position:absolute;right:-10px;top:0;bottom:0;width:55%;background:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 400 200'%3E%3Cpath d='M0,80 C70,40 140,140 220,90 C300,40 360,130 400,90 L400,200 L0,200 Z' fill='rgba(147,197,253,0.35)'/%3E%3C/svg%3E") no-repeat right center/cover;pointer-events:none;}
.wave-vehicle{font-size:68px;flex-shrink:0;}
.wave-info h3{font-size:19px;font-weight:900;color:#1e3a8a;margin-bottom:5px;}
.wave-info p{font-size:12px;color:#3b82f6;font-weight:700;}
.info-grid-icon{display:grid;grid-template-columns:repeat(3,1fr);}
.igf{padding:14px 18px;border-right:1px solid var(--border);border-top:1px solid var(--border);display:flex;align-items:flex-start;gap:12px;}
.igf:nth-child(3n){border-right:none;}
.igf-icon{width:32px;height:32px;border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:14px;flex-shrink:0;margin-top:2px;}
.igf-label{font-size:9px;font-weight:800;letter-spacing:1.5px;color:var(--muted);text-transform:uppercase;margin-bottom:4px;}
.igf-value{font-size:13px;font-weight:800;color:var(--text);}
.charts-row{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:18px;}
.chart-card{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:18px;}
.chart-title{font-size:12px;font-weight:800;color:var(--text);margin-bottom:12px;display:flex;align-items:center;gap:7px;}
canvas{max-height:200px !important;}
.table-card{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:18px;margin-bottom:18px;}
.table-title{font-size:12px;font-weight:800;color:var(--text);margin-bottom:12px;}
table{width:100%;border-collapse:collapse;}
thead th{font-size:10px;font-weight:800;letter-spacing:.5px;color:var(--muted);text-align:left;padding:9px 12px;border-bottom:1px solid var(--border);background:#f8fafc;text-transform:uppercase;}
tbody td{padding:10px 12px;font-size:12px;border-bottom:1px solid #f1f5f9;color:var(--text);font-weight:600;}
tbody tr:last-child td{border-bottom:none;}
tbody tr:hover td{background:#f8fafc;}
.badge{display:inline-flex;align-items:center;gap:3px;border-radius:20px;padding:3px 9px;font-size:10px;font-weight:800;}
.badge-success{background:#f0fdf4;color:#16a34a;border:1px solid #bbf7d0;}
.badge-warn{background:#fffbeb;color:#d97706;border:1px solid #fed7aa;}
.badge-blue{background:#eff6ff;color:#3b82f6;border:1px solid #bfdbfe;}
.plate-chip{font-family:monospace;font-size:11px;font-weight:800;color:#1d4ed8;background:#eff6ff;padding:2px 7px;border-radius:5px;}
.no-case{display:flex;flex-direction:column;align-items:center;justify-content:center;padding:50px;text-align:center;background:var(--surface);border:1px solid var(--border);border-radius:14px;}
.no-case-icon{font-size:44px;margin-bottom:14px;opacity:.2;}
.no-case-text{font-size:14px;font-weight:700;color:var(--muted);}
.vt-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:18px;}
.vt-card{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:14px;text-align:center;}
.vt-icon{font-size:26px;margin-bottom:7px;}.vt-type{font-size:11px;font-weight:700;color:var(--text);}.vt-count{font-size:20px;font-weight:900;color:#3b82f6;margin-top:2px;}
.cctv-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;}
.cctv-card{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:34px 18px;text-align:center;}
.cctv-icon{font-size:32px;opacity:.25;margin-bottom:8px;}.cctv-name{font-size:12px;font-weight:700;color:var(--muted);}.cctv-status{font-size:11px;font-weight:800;margin-top:6px;}
.s-online{color:#22c55e;}.s-offline{color:#ef4444;}.s-degraded{color:#f59e0b;}
.toast{position:fixed;top:70px;right:20px;background:white;border:1px solid #fecaca;border-left:4px solid #ef4444;border-radius:12px;padding:12px 16px;font-size:12px;font-weight:700;color:var(--text);z-index:9990;box-shadow:0 10px 40px rgba(0,0,0,.12);display:none;max-width:340px;}
.toast.show{display:flex;align-items:center;gap:8px;animation:slideIn .3s ease;}
@keyframes slideIn{from{opacity:0;transform:translateX(20px);}to{opacity:1;transform:translateX(0);}}
@keyframes flash{0%,100%{opacity:1;}50%{opacity:.5;}}
@keyframes pulse{0%,100%{opacity:1;transform:scale(1);}50%{opacity:.5;transform:scale(.8);}}
</style>
<style>""" + BELL_CSS + """</style>
</head>
<body>
<div class="toast" id="toast"><span>🚨</span><span id="toast-msg"></span></div>

<!-- NOTIFICATION PANEL -->
<div class="notif-panel" id="notifPanel">
  <div class="notif-hdr">
    <div class="notif-hdr-title">🔔 Notifications</div>
    <button class="notif-mark" onclick="_markAllRead()">Mark all read</button>
  </div>
  <div class="notif-scroll" id="notifList"></div>
  <div class="notif-ftr"><a href="#" onclick="_closeNotif();showSection('analytics',null);return false;">View all activity →</a></div>
</div>

<div class="sidebar">
  <div class="sb-logo">
    <div class="sb-logo-row">
      <div class="sb-badge">⭐</div>
      <div><div class="sb-name">AcciSense</div><div class="sb-sub">Police Portal</div></div>
    </div>
  </div>
  <div class="sb-section">Main</div>
  <div class="nav active" onclick="showSection('dashboard',this)"><span class="nav-ic">🏠</span> Overview</div>
  <div class="nav" onclick="showSection('live',this)"><span class="nav-ic">🚨</span> Incoming Case<span class="nav-badge" id="live-badge" style="display:none">1</span></div>
  <div class="nav" onclick="showSection('analytics',this)"><span class="nav-ic">📈</span> Analytics</div>
  <div class="nav" onclick="showSection('dispatch',this)"><span class="nav-ic">🏆</span> Dispatch Rankings</div>
  <div class="sb-section">Police</div>
  <div class="nav" onclick="showSection('patrol',this)"><span class="nav-ic">🚓</span> Patrol Status</div>
  <div class="nav" onclick="showSection('vehicles',this)"><span class="nav-ic">🚗</span> Vehicles</div>
  <div class="nav" onclick="showSection('officers',this)"><span class="nav-ic">👮</span> Officers On Duty</div>
  <div class="nav" onclick="showSection('cctv',this)"><span class="nav-ic">📷</span> Crime Monitor</div>
  <div class="sb-bottom">
    <div class="sb-user">
      <div class="av">⭐</div>
      <div><div class="un">AcciSense</div><div class="ur">Unit ID: PLEU-001</div></div>
    </div>
  </div>
</div>

<div class="topbar">
  <div class="tb-title">Police Emergency Center</div>
  <div class="tb-right">
    <div class="live-pill"><div class="lp-dot"></div>LIVE</div>
    <div class="tb-time" id="clock">--:--:--</div>
    <div class="bell-wrap" id="bellWrap" onclick="_toggleNotif(event)">
      <button class="bell-btn" title="Notifications">🔔</button>
      <div class="bell-badge hidden" id="bellBadge">0</div>
    </div>
    <a href="/" class="tb-logout">Logout</a>
  </div>
</div>

<div class="main">
  <div id="section-dashboard">
    <div class="page-title">Welcome back, Emergency Unit</div>
    <div class="page-sub">Monitor incoming emergency cases and manage police dispatch response</div>
    <div class="persist-info"><div class="persist-left"><div class="persist-dot"></div><span id="persist-label">Loading...</span></div><span>📅</span></div>
    <div class="stats-row">
      <div class="stat-card"><div class="stat-top"><div class="stat-label">Accepted (All Time)</div><div class="stat-icon-box" style="background:#f0fdf4;color:#16a34a;font-size:20px;">✔</div></div><div class="stat-value" id="st-accepted">0</div><div class="stat-sub">Across all sessions</div></div>
      <div class="stat-card"><div class="stat-top"><div class="stat-label">Declined (All Time)</div><div class="stat-icon-box" style="background:#fef2f2;color:#ef4444;font-size:20px;">✖</div></div><div class="stat-value" id="st-declined" style="color:#ef4444">0</div><div class="stat-sub">Passed to next station</div></div>
      <div class="stat-card"><div class="stat-top"><div class="stat-label">Avg Response</div><div class="stat-icon-box" style="background:#fffbeb;color:#d97706">⏱</div></div><div class="stat-value" id="st-avgtime" style="color:#d97706">--</div><div class="stat-sub">Seconds</div></div>
      <div class="stat-card"><div class="stat-top"><div class="stat-label">Active Incidents</div><div class="stat-icon-box" style="background:#fef2f2;color:#ef4444">🚨</div></div><div class="stat-value" id="st-active" style="color:#ef4444">0</div><div class="stat-sub">Requiring action</div></div>
    </div>
    <div id="dash-alert"></div>
    <div class="charts-row">
      <div class="chart-card"><div class="chart-title">📈 Incident Trend</div><canvas id="trendChart"></canvas></div>
      <div class="chart-card"><div class="chart-title">🚗 Vehicle Types</div><canvas id="vtChart"></canvas></div>
    </div>
    <div class="table-card">
      <div class="table-title">📋 Recent Incidents</div>
      <table><thead><tr><th>Date</th><th>Time</th><th>Vehicle 1</th><th>Vehicle 2</th><th>Type</th><th>Police Station</th><th>Hospital</th><th>Status</th></tr></thead>
      <tbody id="dash-tbody"><tr><td colspan="8" style="text-align:center;color:var(--muted);padding:28px;">No incidents recorded yet</td></tr></tbody></table>
    </div>
  </div>
  <div id="section-live" style="display:none">
    <div class="page-title">Live Incident</div><div class="page-sub">Real-time accident monitoring</div>
    <div id="live-detail"></div>
  </div>
  <div id="section-analytics" style="display:none">
    <div class="page-title">Analytics</div><div class="page-sub">Incident patterns and performance metrics</div>
    <div class="charts-row">
      <div class="chart-card"><div class="chart-title">🚓 Cases by Police Station</div><canvas id="stationChart"></canvas></div>
      <div class="chart-card"><div class="chart-title">⏱ Response Time Distribution</div><canvas id="responseChart"></canvas></div>
    </div>
    <div class="charts-row">
      <div class="chart-card"><div class="chart-title">📅 Cases by Date</div><canvas id="weeklyChart"></canvas></div>
      <div class="chart-card"><div class="chart-title">🕐 Incidents by Hour</div><canvas id="hourChart"></canvas></div>
    </div>
  </div>
  <div id="section-dispatch" style="display:none">
    <div class="page-title">Dispatch Rankings</div><div class="page-sub">Performance-based station leaderboard</div>
    <div class="table-card"><table><thead><tr><th>Rank</th><th>Police Station</th><th>Cases Handled</th><th>Status</th></tr></thead><tbody id="dispatch-tbody"><tr><td colspan="4" style="text-align:center;color:var(--muted);padding:28px;">No data</td></tr></tbody></table></div>
  </div>
  <div id="section-patrol" style="display:none">
    <div class="page-title">Patrol Status</div><div class="page-sub">Active patrol units</div>
    <div class="charts-row">
      <div class="chart-card" style="text-align:center;padding:38px;"><div style="font-size:34px;opacity:.3;margin-bottom:8px;">🚓</div><div style="font-size:12px;font-weight:700;color:var(--muted);">Patrol Unit Alpha — Zone 3</div><div style="font-size:11px;color:#22c55e;font-weight:800;margin-top:5px;">● On Duty</div></div>
      <div class="chart-card" style="text-align:center;padding:38px;"><div style="font-size:34px;opacity:.3;margin-bottom:8px;">🚓</div><div style="font-size:12px;font-weight:700;color:var(--muted);">Patrol Unit Beta — Zone 1</div><div style="font-size:11px;color:#22c55e;font-weight:800;margin-top:5px;">● On Duty</div></div>
    </div>
  </div>
  <div id="section-vehicles" style="display:none">
    <div class="page-title">Vehicle Type Analysis</div><div class="page-sub">Breakdown of vehicle types in accidents</div>
    <div class="vt-grid" id="vt-grid"></div>
    <div class="charts-row">
      <div class="chart-card"><div class="chart-title">🍩 Type Distribution</div><canvas id="vtDoughnut"></canvas></div>
      <div class="chart-card"><div class="chart-title">📊 Incidents by Type</div><canvas id="vtBar"></canvas></div>
    </div>
    <div class="table-card"><div class="table-title">🚗 Vehicle Incident Log</div>
      <table><thead><tr><th>Date</th><th>Time</th><th>Plate</th><th>Type</th><th>Icon</th><th>Hospital</th><th>Status</th></tr></thead><tbody id="vt-tbody"><tr><td colspan="7" style="text-align:center;color:var(--muted);padding:28px;">No data</td></tr></tbody></table>
    </div>
  </div>
  <div id="section-officers" style="display:none">
    <div class="page-title">Officers On Duty</div><div class="page-sub">Current shift</div>
    <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:14px;">
      <div class="chart-card" style="display:flex;align-items:center;gap:12px;padding:18px;"><div style="font-size:30px;">👮</div><div><div style="font-weight:800;font-size:13px;">Sgt. Williams</div><div style="font-size:11px;color:var(--muted);">Traffic Division</div><div style="font-size:11px;color:#22c55e;font-weight:800;margin-top:4px;">● On Duty</div></div></div>
      <div class="chart-card" style="display:flex;align-items:center;gap:12px;padding:18px;"><div style="font-size:30px;">👮</div><div><div style="font-weight:800;font-size:13px;">Officer A. Lee</div><div style="font-size:11px;color:var(--muted);">Patrol Unit</div><div style="font-size:11px;color:#22c55e;font-weight:800;margin-top:4px;">● On Duty</div></div></div>
      <div class="chart-card" style="display:flex;align-items:center;gap:12px;padding:18px;"><div style="font-size:30px;">👮</div><div><div style="font-weight:800;font-size:13px;">Cpl. Rodriguez</div><div style="font-size:11px;color:var(--muted);">Highway Unit</div><div style="font-size:11px;color:#f59e0b;font-weight:800;margin-top:4px;">⚠ Break</div></div></div>
    </div>
  </div>
  <div id="section-cctv" style="display:none">
    <div class="page-title">Crime Monitor</div><div class="page-sub">Surveillance cameras</div>
    <div class="cctv-grid">
      <div class="cctv-card"><div class="cctv-icon">📷</div><div class="cctv-name">Camera 01 — NH-44 Junction</div><div class="cctv-status s-online">● Online</div></div>
      <div class="cctv-card"><div class="cctv-icon">📷</div><div class="cctv-name">Camera 02 — City Mall Signal</div><div class="cctv-status s-online">● Online</div></div>
      <div class="cctv-card"><div class="cctv-icon">📷</div><div class="cctv-name">Camera 03 — Railway Gate</div><div class="cctv-status s-offline">● Offline</div></div>
      <div class="cctv-card"><div class="cctv-icon">📷</div><div class="cctv-name">Camera 04 — Bus Terminal</div><div class="cctv-status s-online">● Online</div></div>
      <div class="cctv-card"><div class="cctv-icon">📷</div><div class="cctv-name">Camera 05 — Airport Road</div><div class="cctv-status s-online">● Online</div></div>
      <div class="cctv-card"><div class="cctv-icon">📷</div><div class="cctv-name">Camera 06 — Industrial Zone</div><div class="cctv-status s-degraded">● Degraded</div></div>
    </div>
  </div>
</div>

<script>
""" + BELL_JS + """
var CS={},lastCases=0;
var trendC,vtC,stationC,responseC,weeklyC,hourC,vtDC,vtBC;
var vtCounts={},stationCounts={},hourCounts=new Array(24).fill(0),responseTimes=[],dateGroupCounts={};

setInterval(function(){document.getElementById('clock').textContent=new Date().toLocaleTimeString('en-IN',{hour12:false});},1000);

function showSection(name,el){
  document.querySelectorAll('[id^="section-"]').forEach(function(s){s.style.display='none';});
  document.getElementById('section-'+name).style.display='block';
  document.querySelectorAll('.nav').forEach(function(n){n.classList.remove('active');});
  if(el)el.classList.add('active');
  if(name==='analytics')initAnalyticsCharts();
  if(name==='vehicles')renderVehicles();
  if(name==='dispatch')renderDispatch();
  _closeNotif();
}

function showToast(msg){var t=document.getElementById('toast');document.getElementById('toast-msg').textContent=msg;t.classList.add('show');setTimeout(function(){t.classList.remove('show');},6000);}

var VT={TN01:{type:"Sedan",icon:"🚗"},TN02:{type:"SUV",icon:"🚙"},TN03:{type:"Truck",icon:"🚛"},MH:{type:"Motorcycle",icon:"🏍️"},KA:{type:"Auto-Rickshaw",icon:"🛺"},AP:{type:"Bus",icon:"🚌"},TS:{type:"Van",icon:"🚐"},DL:{type:"Hatchback",icon:"🚘"},GJ:{type:"Pickup",icon:"🛻"},RJ:{type:"Minibus",icon:"🚎"},HR:{type:"Bicycle",icon:"🚲"},UP:{type:"Ambulance",icon:"🚑"},WB:{type:"Taxi",icon:"🚕"},MP:{type:"Police Car",icon:"🚓"},PB:{type:"Fire Truck",icon:"🚒"}};
function getVI(p){for(var k in VT){if(p.toUpperCase().startsWith(k))return VT[k];}return{type:"Vehicle",icon:"🚗"};}
function fmtTime(iso){return iso?new Date(iso).toLocaleTimeString('en-IN',{hour12:false}):'--';}

function buildPoliceCard(s){
  if(!s.current_case)return'<div class="no-case"><div class="no-case-icon">🛡️</div><div class="no-case-text">No active incident — System monitoring</div></div>';
  var c=s.current_case,v1=getVI(c.vehicle_1),v2=getVI(c.vehicle_2);
  var isPending=s.case_status==='pending',isDispatched=s.case_status==='ambulance_dispatched';
  var pill=isPending?'<div class="disp-pill-emergency">🚨 EMERGENCY ALERT</div>':'<div class="disp-pill-resolved">✅ RESOLVED</div>';
  var badge=isPending?'<div class="disp-status-badge dsb-red">PENDING</div>':'<div class="disp-status-badge dsb-blue">POLICE DISPATCHED</div>';
  var at=fmtTime(s.alert_time),act=fmtTime(s.accept_time);
  var hero=isDispatched?'<div class="wave-hero blue-wave"><div class="wave-vehicle">🚓</div><div class="wave-info"><h3>Police Dispatched</h3><p>Alert: '+at+' &nbsp;|&nbsp; Accepted: '+act+'</p></div></div>':isPending?'<div class="wave-hero red-wave"><div class="wave-vehicle">🚨</div><div class="wave-info"><h3 style="color:#7f1d1d;">Emergency Alert</h3><p style="color:#dc2626;">Received at '+at+' — Awaiting dispatch</p></div></div>':'';
  return'<div class="disp-card-wrap '+(isPending?'emergency':'resolved')+'">'
    +'<div class="disp-header">'+pill+'<div class="disp-title">Emergency Dispatch Response Required</div>'+badge+'</div>'
    +hero
    +'<div class="info-grid-icon">'
    +'<div class="igf"><div class="igf-icon" style="background:#eff6ff">🏛</div><div><div class="igf-label">Station Contacted</div><div class="igf-value">'+(s.selected_police||'N/A')+'</div></div></div>'
    +'<div class="igf"><div class="igf-icon" style="background:#f0fdf4">📍</div><div><div class="igf-label">Distance</div><div class="igf-value">6.2 km</div></div></div>'
    +'<div class="igf"><div class="igf-icon" style="background:#fffbeb">⏰</div><div><div class="igf-label">Alert Time</div><div class="igf-value">'+at+'</div></div></div>'
    +'<div class="igf"><div class="igf-icon" style="background:#eff6ff">🚗</div><div><div class="igf-label">Vehicles</div><div class="igf-value">'+v1.icon+' '+c.vehicle_1+' / '+v2.icon+' '+c.vehicle_2+'</div></div></div>'
    +'<div class="igf"><div class="igf-icon" style="background:#f0fdf4">🌐</div><div><div class="igf-label">GPS Location</div><div class="igf-value">'+(c.latitude||'N/A')+', '+(c.longitude||'N/A')+'</div></div></div>'
    +'<div class="igf"><div class="igf-icon" style="background:#faf5ff">👮</div><div><div class="igf-label">Officers Assigned</div><div class="igf-value">Sgt. Williams, A. Lee</div></div></div>'
    +'</div></div>';
}

function renderDispatch(){
  var tbody=document.getElementById('dispatch-tbody');
  var entries=Object.entries(stationCounts).sort(function(a,b){return b[1]-a[1];});
  if(!entries.length){tbody.innerHTML='<tr><td colspan="4" style="text-align:center;color:var(--muted);padding:28px;">No data</td></tr>';return;}
  var medals=['🥇','🥈','🥉'];
  tbody.innerHTML=entries.map(function(e,i){return'<tr><td>'+( medals[i]||i+1)+'</td><td style="font-weight:800">'+e[0]+'</td><td><strong>'+e[1]+'</strong></td><td><span class="badge badge-success">Active</span></td></tr>';}).join('');
}

function renderVehicles(){
  var grid=document.getElementById('vt-grid');
  if(!Object.keys(vtCounts).length){grid.innerHTML='<div style="grid-column:1/-1;text-align:center;color:var(--muted);padding:38px;font-size:14px;font-weight:700;">No vehicle data yet</div>';return;}
  grid.innerHTML=Object.entries(vtCounts).map(function(e){var icon=Object.values(VT).find(function(v){return v.type===e[0];});return'<div class="vt-card"><div class="vt-icon">'+(icon?icon.icon:'🚗')+'</div><div class="vt-type">'+e[0]+'</div><div class="vt-count">'+e[1]+'</div></div>';}).join('');
  var keys=Object.keys(vtCounts),vals=Object.values(vtCounts);
  if(!vtDC){
    vtDC=new Chart(document.getElementById('vtDoughnut'),{type:'doughnut',data:{labels:keys,datasets:[{data:vals,backgroundColor:['#3b82f6','#22d3a5','#f59e0b','#ef4444','#a78bfa','#34d399','#fb7185','#60a5fa'],borderWidth:0}]},options:{plugins:{legend:{position:'bottom',labels:{boxWidth:10,padding:8,font:{size:11,family:'Nunito'}}}},cutout:'60%',animation:false}});
    vtBC=new Chart(document.getElementById('vtBar'),{type:'bar',data:{labels:keys,datasets:[{label:'Count',data:vals,backgroundColor:'rgba(59,130,246,0.6)',borderColor:'#3b82f6',borderWidth:1,borderRadius:6}]},options:{plugins:{legend:{display:false}},scales:{y:{beginAtZero:true},x:{grid:{display:false}}},animation:false}});
  }else{vtDC.data.labels=keys;vtDC.data.datasets[0].data=vals;vtDC.update();vtBC.data.labels=keys;vtBC.data.datasets[0].data=vals;vtBC.update();}
  var tbody=document.getElementById('vt-tbody');
  if(CS.case_history&&CS.case_history.length){
    tbody.innerHTML=CS.case_history.flatMap(function(c){return[{plate:c.v1,type:c.v1type||getVI(c.v1).type,icon:c.v1icon||getVI(c.v1).icon,date:c.date,time:c.time,hospital:c.hospital,status:c.status},{plate:c.v2,type:c.v2type||getVI(c.v2).type,icon:c.v2icon||getVI(c.v2).icon,date:c.date,time:c.time,hospital:c.hospital,status:c.status}];}).reverse().map(function(r){return'<tr><td>'+r.date+'</td><td>'+r.time+'</td><td><span class="plate-chip">'+r.plate+'</span></td><td>'+r.type+'</td><td style="font-size:16px">'+r.icon+'</td><td>'+r.hospital+'</td><td><span class="badge badge-success">'+r.status+'</span></td></tr>';}).join('');
  }
}

function initCharts(){
  Chart.defaults.font.family='Nunito';Chart.defaults.font.weight='700';Chart.defaults.color='#64748b';
  trendC=new Chart(document.getElementById('trendChart'),{type:'line',data:{labels:[],datasets:[{label:'Incidents',data:[],borderColor:'#3b82f6',backgroundColor:'rgba(59,130,246,0.08)',tension:.4,fill:true,pointBackgroundColor:'#3b82f6',pointRadius:4,borderWidth:2}]},options:{plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,grid:{color:'rgba(0,0,0,0.04)'}},x:{grid:{display:false},ticks:{maxRotation:45,font:{size:9}}}},animation:false}});
  vtC=new Chart(document.getElementById('vtChart'),{type:'doughnut',data:{labels:[],datasets:[{data:[],backgroundColor:['#3b82f6','#22d3a5','#f59e0b','#ef4444','#a78bfa','#34d399','#fb7185','#60a5fa'],borderWidth:0}]},options:{plugins:{legend:{position:'bottom',labels:{boxWidth:10,padding:8,font:{size:11}}}},cutout:'60%',animation:false}});
}

function initAnalyticsCharts(){
  if(stationC)return;
  var labels=Object.keys(stationCounts).length?Object.keys(stationCounts):['No Data'];
  var vals=Object.values(stationCounts).length?Object.values(stationCounts):[0];
  stationC=new Chart(document.getElementById('stationChart'),{type:'bar',data:{labels:labels,datasets:[{label:'Cases',data:vals,backgroundColor:'rgba(59,130,246,0.6)',borderColor:'#3b82f6',borderWidth:1,borderRadius:6}]},options:{plugins:{legend:{display:false}},scales:{y:{beginAtZero:true},x:{grid:{display:false},ticks:{maxRotation:45,font:{size:9}}}},animation:false}});
  responseC=new Chart(document.getElementById('responseChart'),{type:'bar',data:{labels:['0-30s','31-60s','61-120s','120s+'],datasets:[{label:'Cases',data:[0,0,0,0],backgroundColor:['rgba(34,211,165,0.6)','rgba(59,130,246,0.6)','rgba(245,158,11,0.6)','rgba(239,68,68,0.6)'],borderWidth:0,borderRadius:6}]},options:{plugins:{legend:{display:false}},scales:{y:{beginAtZero:true},x:{grid:{display:false}}},animation:false}});
  var dl=Object.keys(dateGroupCounts).sort();
  weeklyC=new Chart(document.getElementById('weeklyChart'),{type:'line',data:{labels:dl.length?dl:['No Data'],datasets:[{label:'Incidents',data:dl.length?dl.map(function(d){return dateGroupCounts[d];}):[ 0],borderColor:'#ef4444',backgroundColor:'rgba(239,68,68,0.05)',tension:.4,fill:true,borderWidth:2}]},options:{plugins:{legend:{display:false}},scales:{y:{beginAtZero:true},x:{grid:{display:false},ticks:{maxRotation:45,font:{size:9}}}},animation:false}});
  hourC=new Chart(document.getElementById('hourChart'),{type:'bar',data:{labels:Array.from({length:24},function(_,i){return i+':00';}),datasets:[{label:'Incidents',data:hourCounts,backgroundColor:'rgba(167,139,250,0.6)',borderColor:'#a78bfa',borderWidth:1,borderRadius:4}]},options:{plugins:{legend:{display:false}},scales:{y:{beginAtZero:true},x:{grid:{display:false},ticks:{maxRotation:90,font:{size:9}}}},animation:false}});
}

async function poll(){
  try{
    var r=await fetch('/api/state');var s=await r.json();CS=s;
    var hist=s.case_history||[];
    var earliest=hist.length?hist[0].date:'—',latest=hist.length?hist[hist.length-1].date:'—';
    document.getElementById('persist-label').textContent=hist.length?'Persistent data loaded — '+hist.length+' total cases from '+earliest+' to '+latest:'Persistent storage active — No cases yet recorded';
    var hs=s.hospital_stats||{},totA=0,totD=0,totT=0,totN=0;
    Object.values(hs).forEach(function(h){totA+=h.accepted;totD+=h.declined;h.times.forEach(function(t){totT+=t;totN++;});});
    document.getElementById('st-accepted').textContent=totA;
    document.getElementById('st-declined').textContent=totD;
    document.getElementById('st-avgtime').textContent=totN>0?Math.round(totT/totN):'--';
    document.getElementById('st-active').textContent=s.case_status==='pending'?1:0;
    var lb=document.getElementById('live-badge');lb.style.display=s.case_status==='pending'?'inline':'none';
    if(s.current_case&&s.total_cases>lastCases&&s.case_status==='pending'){
      var v1=getVI(s.current_case.vehicle_1),v2=getVI(s.current_case.vehicle_2);
      showToast('🚨 Accident: '+v1.icon+s.current_case.vehicle_1+' & '+v2.icon+s.current_case.vehicle_2);
      _pushNotif('🚨','#fef2f2','New Emergency Alert','Accident involving '+s.current_case.vehicle_1+' & '+s.current_case.vehicle_2+' detected.');
    }
    if(s.total_cases>0)lastCases=s.total_cases;
    var cardHtml=buildPoliceCard(s);
    document.getElementById('dash-alert').innerHTML=cardHtml;
    document.getElementById('live-detail').innerHTML=cardHtml;
    vtCounts={};stationCounts={};hourCounts=new Array(24).fill(0);responseTimes=[];dateGroupCounts={};
    hist.forEach(function(c){
      [c.v1,c.v2].forEach(function(p){var vt=getVI(p).type;vtCounts[vt]=(vtCounts[vt]||0)+1;});
      stationCounts[c.police]=(stationCounts[c.police]||0)+1;
      var h=parseInt((c.time||'').split(':')[0]);if(!isNaN(h))hourCounts[h]++;
      responseTimes.push(c.response_time);
      var d=c.date||'Unknown';dateGroupCounts[d]=(dateGroupCounts[d]||0)+1;
    });
    var tbody=document.getElementById('dash-tbody');
    if(hist.length){tbody.innerHTML=hist.slice().reverse().slice(0,10).map(function(c){return'<tr><td>'+(c.date||'--')+'</td><td>'+c.time+'</td><td><span class="plate-chip">'+c.v1+'</span></td><td><span class="plate-chip">'+c.v2+'</span></td><td>'+(c.v1icon||getVI(c.v1).icon)+' '+(c.v1type||getVI(c.v1).type)+'</td><td style="font-size:11px;color:#f59e0b;font-weight:700;">'+c.police+'</td><td>'+c.hospital+'</td><td><span class="badge badge-success">'+c.status+'</span></td></tr>';}).join('');}
    var dl=Object.keys(dateGroupCounts).sort();
    trendC.data.labels=dl.length?dl:['No data'];trendC.data.datasets[0].data=dl.length?dl.map(function(d){return dateGroupCounts[d];}):[ 0];trendC.update();
    var vl=Object.keys(vtCounts),vv=Object.values(vtCounts);vtC.data.labels=vl.length?vl:['No Data'];vtC.data.datasets[0].data=vv.length?vv:[1];vtC.update();
    if(document.getElementById('section-analytics').style.display!=='none'&&stationC){
      stationC.data.labels=Object.keys(stationCounts)||['No Data'];stationC.data.datasets[0].data=Object.values(stationCounts)||[0];stationC.update();
      var bins=[0,0,0,0];responseTimes.forEach(function(t){if(t<=30)bins[0]++;else if(t<=60)bins[1]++;else if(t<=120)bins[2]++;else bins[3]++;});responseC.data.datasets[0].data=bins;responseC.update();
      hourC.data.datasets[0].data=hourCounts;hourC.update();
      weeklyC.data.labels=dl.length?dl:['No Data'];weeklyC.data.datasets[0].data=dl.length?dl.map(function(d){return dateGroupCounts[d];}):[ 0];weeklyC.update();
    }
  }catch(e){console.error(e);}
}

// Init bell with police notifications
_initBell('#3b82f6','#eff6ff',[
  {icon:'🚨',bg:'#fef2f2',title:'System Online',desc:'Police dashboard connected. Monitoring live incidents.',time:'Just now',unread:true},
  {icon:'🚓',bg:'#eff6ff',title:'Patrol Units Active',desc:'All patrol units are reporting in from assigned zones.',time:'5 min ago',unread:true},
  {icon:'📋',bg:'#f0fdf4',title:'Shift Started',desc:'Day shift handover complete. All logs reviewed.',time:'1 hr ago',unread:false},
  {icon:'📷',bg:'#fffbeb',title:'Camera 03 Offline',desc:'Railway Gate camera has gone offline. Maintenance notified.',time:'2 hr ago',unread:false},
]);
initCharts();
poll();
setInterval(poll,2000);
</script>
</body>
</html>"""

# ─────────────────────────────────────────────
# HOSPITAL DASHBOARD — BLUE THEME (matching screenshot)
# ─────────────────────────────────────────────
HOSPITAL_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>AcciSense — Hospital Dashboard</title>
<link href="https://fonts.googleapis.com/css2?family=Nunito:wght@400;500;600;700;800;900&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<style>
*{margin:0;padding:0;box-sizing:border-box;}
:root{
  --bg:#e8f0fb;
  --sidebar:#0a1628;
  --surface:#fff;
  --border:#dce6f5;
  --accent:#2563eb;
  --accent2:#3b82f6;
  --warn:#f59e0b;
  --danger:#ef4444;
  --text:#0f172a;
  --muted:#64748b;
  --sw:195px;
  --notif-unread:#eff6ff;
}
body{background:var(--bg);font-family:'Nunito',sans-serif;color:var(--text);display:flex;min-height:100vh;}

/* ── SIDEBAR ── */
.sidebar{width:var(--sw);background:var(--sidebar);display:flex;flex-direction:column;position:fixed;top:0;left:0;height:100vh;z-index:100;overflow-y:auto;}
.sb-logo{padding:20px 16px 16px;border-bottom:1px solid rgba(255,255,255,.07);}
.sb-logo-row{display:flex;align-items:center;gap:10px;}
.sb-badge{width:42px;height:42px;background:linear-gradient(135deg,#2563eb,#1e40af);border-radius:12px;display:flex;align-items:center;justify-content:center;font-size:20px;flex-shrink:0;box-shadow:0 4px 12px rgba(37,99,235,.5);}
.sb-name{font-size:14px;font-weight:900;color:white;}.sb-sub{font-size:9px;color:#334d7a;letter-spacing:1.5px;margin-top:1px;text-transform:uppercase;}
.sb-section{padding:16px 16px 5px;font-size:9px;font-weight:800;letter-spacing:2px;color:#1e3a5f;text-transform:uppercase;}
.nav{display:flex;align-items:center;gap:9px;padding:9px 14px 9px 16px;cursor:pointer;transition:all .2s;font-size:12px;font-weight:700;color:#3a6490;margin:1px 0;border-right:3px solid transparent;}
.nav:hover{background:rgba(59,130,246,.08);color:#60a5fa;}
.nav.active{background:rgba(59,130,246,.18);color:#60a5fa;border-right-color:#3b82f6;}
.nav-ic{font-size:13px;width:17px;text-align:center;}
.nav-badge{margin-left:auto;background:#ef4444;color:white;border-radius:20px;padding:1px 6px;font-size:9px;font-weight:800;}
.sb-bottom{margin-top:auto;padding:14px;border-top:1px solid rgba(255,255,255,.06);}
.sb-user{display:flex;align-items:center;gap:9px;}
.av{width:34px;height:34px;background:linear-gradient(135deg,#2563eb,#1e40af);border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:16px;color:white;}
.un{font-size:11px;font-weight:800;color:#e2e8f0;}.ur{font-size:9px;color:#334d7a;letter-spacing:1px;text-transform:uppercase;}

/* ── TOPBAR ── */
.topbar{position:fixed;top:0;left:var(--sw);right:0;height:58px;background:rgba(255,255,255,0.92);backdrop-filter:blur(12px);border-bottom:1px solid var(--border);display:flex;align-items:center;padding:0 24px;z-index:50;box-shadow:0 1px 8px rgba(37,99,235,.08);}
.tb-title{font-size:14px;font-weight:900;color:var(--text);}
.tb-right{margin-left:auto;display:flex;align-items:center;gap:12px;}
.live-pill{display:flex;align-items:center;gap:6px;background:#eff6ff;border:1px solid #bfdbfe;border-radius:20px;padding:5px 12px;font-size:11px;font-weight:800;color:#2563eb;letter-spacing:1px;}
.lp-dot{width:6px;height:6px;border-radius:50%;background:#3b82f6;animation:pulse 1.2s infinite;}
.tb-time{font-size:12px;color:var(--muted);font-weight:700;}
.tb-logout{background:#f8fafc;border:1px solid var(--border);color:var(--muted);padding:6px 12px;border-radius:8px;cursor:pointer;font-size:12px;font-weight:700;text-decoration:none;}

/* ── MAIN ── */
.main{margin-left:var(--sw);margin-top:58px;flex:1;padding:24px;overflow-y:auto;}
.page-title{font-size:20px;font-weight:900;color:var(--text);margin-bottom:3px;}
.page-sub{font-size:13px;color:var(--muted);margin-bottom:20px;font-weight:600;}

/* ── PERSIST BAR ── */
.persist-info{display:flex;align-items:center;justify-content:space-between;gap:10px;background:#eff6ff;border:1px solid #bfdbfe;border-radius:10px;padding:10px 16px;margin-bottom:18px;font-size:12px;color:#1e40af;font-weight:700;}
.persist-left{display:flex;align-items:center;gap:8px;}
.persist-dot{width:8px;height:8px;border-radius:50%;background:#3b82f6;flex-shrink:0;}

/* ── STAT CARDS ── */
.stats-row{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:20px;}
.stat-card{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:18px;transition:box-shadow .2s;box-shadow:0 2px 8px rgba(37,99,235,.04);}
.stat-card:hover{box-shadow:0 8px 24px rgba(37,99,235,.10);}
.stat-top{display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:10px;}
.stat-label{font-size:11px;font-weight:700;color:var(--muted);}
.stat-icon-box{width:36px;height:36px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:16px;}
.stat-value{font-size:28px;font-weight:900;color:var(--text);line-height:1;}
.stat-sub{font-size:11px;color:var(--muted);margin-top:4px;font-weight:600;}

/* ── DISPATCH CARD ── */
.disp-card-wrap{border-radius:16px;overflow:hidden;margin-bottom:20px;border:1.5px solid var(--border);background:var(--surface);box-shadow:0 4px 20px rgba(37,99,235,.07);}
.disp-card-wrap.emergency{border-color:#fca5a5;}.disp-card-wrap.resolved{border-color:#93c5fd;}
.disp-header{display:flex;align-items:center;gap:12px;padding:16px 20px;border-bottom:1px solid var(--border);}
.disp-pill-resolved{background:#eff6ff;border:1px solid #93c5fd;color:#1d4ed8;border-radius:8px;padding:5px 12px;font-size:11px;font-weight:800;display:flex;align-items:center;gap:5px;}
.disp-pill-emergency{background:#fef2f2;border:1px solid #fca5a5;color:#dc2626;border-radius:8px;padding:5px 12px;font-size:11px;font-weight:800;animation:flash 1.2s infinite;display:flex;align-items:center;gap:5px;}
.disp-title{font-size:14px;font-weight:800;color:var(--text);}
.disp-status-badge{margin-left:auto;padding:6px 16px;border-radius:20px;font-size:11px;font-weight:800;}
.dsb-green{background:#2563eb;color:white;}.dsb-red{background:#dc2626;color:white;}.dsb-blue{background:#1d4ed8;color:white;}.dsb-orange{background:#d97706;color:white;}

/* ── WAVE HERO ── */
.wave-hero{position:relative;padding:24px 28px;min-height:120px;display:flex;align-items:center;gap:20px;overflow:hidden;}
.wave-hero.blue-wave{background:linear-gradient(135deg,#dbeafe,#bfdbfe,#93c5fd);}
.wave-hero.amber-wave{background:linear-gradient(135deg,#fef3c7,#fde68a,#fcd34d);}
.wave-hero::after{content:'';position:absolute;right:-10px;top:0;bottom:0;width:55%;background:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 400 200'%3E%3Cpath d='M0,80 C70,40 140,140 220,90 C300,40 360,130 400,90 L400,200 L0,200 Z' fill='rgba(147,197,253,0.35)'/%3E%3C/svg%3E") no-repeat right center/cover;pointer-events:none;}
.wave-vehicle{font-size:68px;flex-shrink:0;}
.wave-info h3{font-size:19px;font-weight:900;color:#1e3a8a;margin-bottom:5px;}
.wave-info p{font-size:12px;color:#2563eb;font-weight:700;}

/* ── INFO GRID ── */
.info-grid-icon{display:grid;grid-template-columns:repeat(3,1fr);}
.igf{padding:14px 18px;border-right:1px solid var(--border);border-top:1px solid var(--border);display:flex;align-items:flex-start;gap:12px;}
.igf:nth-child(3n){border-right:none;}
.igf-icon{width:32px;height:32px;border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:14px;flex-shrink:0;margin-top:2px;}
.igf-label{font-size:9px;font-weight:800;letter-spacing:1.5px;color:var(--muted);text-transform:uppercase;margin-bottom:4px;}
.igf-value{font-size:13px;font-weight:800;color:var(--text);}

/* ── ACTION BUTTONS ── */
.action-btns{display:flex;gap:14px;padding:16px 20px;background:var(--surface);border-top:1px solid var(--border);}
.btn-accept{flex:1;padding:14px;background:linear-gradient(135deg,#2563eb,#1d4ed8);border:none;border-radius:12px;color:white;font-family:'Nunito',sans-serif;font-size:13px;font-weight:900;cursor:pointer;transition:.25s;box-shadow:0 6px 20px rgba(37,99,235,.35);}
.btn-accept:hover{transform:translateY(-2px);box-shadow:0 10px 28px rgba(37,99,235,.45);}
.btn-decline{flex:1;padding:14px;background:linear-gradient(135deg,#ef4444,#dc2626);border:none;border-radius:12px;color:white;font-family:'Nunito',sans-serif;font-size:13px;font-weight:900;cursor:pointer;transition:.25s;box-shadow:0 6px 20px rgba(239,68,68,.3);}
.btn-decline:hover{transform:translateY(-2px);}

/* ── CHARTS ── */
.charts-row{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:18px;}
.chart-card{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:18px;box-shadow:0 2px 8px rgba(37,99,235,.04);}
.chart-title{font-size:12px;font-weight:800;color:var(--text);margin-bottom:12px;display:flex;align-items:center;gap:7px;}
canvas{max-height:220px !important;}

/* ── TABLES ── */
.table-card{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:18px;margin-bottom:18px;box-shadow:0 2px 8px rgba(37,99,235,.04);}
table{width:100%;border-collapse:collapse;}
thead th{font-size:10px;font-weight:800;letter-spacing:.5px;color:var(--muted);text-align:left;padding:9px 12px;border-bottom:1px solid var(--border);background:#f0f6ff;text-transform:uppercase;}
tbody td{padding:10px 12px;font-size:12px;border-bottom:1px solid #eef2fb;color:var(--text);font-weight:600;}
tbody tr:last-child td{border-bottom:none;}
tbody tr:hover td{background:#f0f6ff;}
.badge{display:inline-flex;align-items:center;gap:3px;border-radius:20px;padding:3px 9px;font-size:10px;font-weight:800;}
.badge-success{background:#eff6ff;color:#2563eb;border:1px solid #bfdbfe;}
.badge-danger{background:#fef2f2;color:#ef4444;border:1px solid #fecaca;}
.badge-warn{background:#fffbeb;color:#d97706;border:1px solid #fed7aa;}
.badge-blue{background:#eff6ff;color:#3b82f6;border:1px solid #bfdbfe;}

/* ── NO CASE ── */
.no-case{display:flex;flex-direction:column;align-items:center;justify-content:center;padding:50px;text-align:center;background:var(--surface);border:1px solid var(--border);border-radius:14px;}
.no-case-icon{font-size:44px;margin-bottom:14px;opacity:.2;}.no-case-text{font-size:14px;font-weight:700;color:var(--muted);}

/* ── BED CARDS ── */
.bed-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:18px;}
.bed-card{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:16px;position:relative;overflow:hidden;box-shadow:0 2px 8px rgba(37,99,235,.04);}
.bed-card::after{content:'';position:absolute;bottom:0;left:0;right:0;height:3px;}
.bed-card.good::after{background:#3b82f6;}.bed-card.warning::after{background:#f59e0b;}.bed-card.critical::after{background:#ef4444;}
.bed-ward{font-size:12px;font-weight:800;color:var(--text);margin-bottom:10px;}
.bed-numbers{display:flex;align-items:baseline;gap:4px;margin-bottom:7px;}
.bed-avail{font-size:28px;font-weight:900;}
.bed-avail.good{color:#3b82f6;}.bed-avail.warning{color:#f59e0b;}.bed-avail.critical{color:#ef4444;}
.bed-total{font-size:12px;color:var(--muted);}
.bed-bar{width:100%;height:5px;background:#e2ecfb;border-radius:3px;overflow:hidden;}
.bed-fill{height:100%;border-radius:3px;}
.bed-fill.good{background:#3b82f6;}.bed-fill.warning{background:#f59e0b;}.bed-fill.critical{background:#ef4444;}
.bed-label{font-size:10px;color:var(--muted);margin-top:5px;font-weight:700;}

/* ── RANKING ── */
.ranking-row{display:flex;align-items:center;gap:12px;padding:12px 0;border-bottom:1px solid #eef2fb;}
.ranking-row:last-child{border-bottom:none;}
.rank-num{font-size:18px;font-weight:900;width:34px;text-align:center;color:var(--muted);}
.rank-name{flex:1;}.rank-name-main{font-size:13px;font-weight:800;color:var(--text);}.rank-name-sub{font-size:11px;color:var(--muted);margin-top:1px;font-weight:600;}
.hosp-stars{color:#f59e0b;font-size:12px;}
.hosp-tag{font-size:10px;padding:3px 9px;border-radius:20px;font-weight:700;}
.tag-fast{background:#eff6ff;color:#2563eb;border:1px solid #bfdbfe;}.tag-slow{background:#fef2f2;color:#ef4444;border:1px solid #fecaca;}

/* ── TOAST ── */
.toast{position:fixed;top:70px;right:20px;background:white;border:1px solid #bfdbfe;border-left:4px solid #2563eb;border-radius:12px;padding:12px 16px;font-size:12px;font-weight:700;color:var(--text);z-index:9990;box-shadow:0 10px 40px rgba(37,99,235,.15);display:none;max-width:340px;}
.toast.show{display:flex;align-items:center;gap:8px;animation:slideIn .3s ease;}

@keyframes slideIn{from{opacity:0;transform:translateX(20px);}to{opacity:1;transform:translateX(0);}}
@keyframes flash{0%,100%{opacity:1;}50%{opacity:.5;}}
@keyframes pulse{0%,100%{opacity:1;transform:scale(1);}50%{opacity:.5;transform:scale(.8);}}
</style>
<style>""" + BELL_CSS + """</style>
</head>
<body>
<div class="toast" id="toast"><span>🏥</span><span id="toast-msg"></span></div>

<!-- NOTIFICATION PANEL -->
<div class="notif-panel" id="notifPanel">
  <div class="notif-hdr">
    <div class="notif-hdr-title">🔔 Notifications</div>
    <button class="notif-mark" onclick="_markAllRead()">Mark all read</button>
  </div>
  <div class="notif-scroll" id="notifList"></div>
  <div class="notif-ftr"><a href="#" onclick="_closeNotif();showSection('analytics',null);return false;">View all activity →</a></div>
</div>

<div class="sidebar">
  <div class="sb-logo">
    <div class="sb-logo-row">
      <div class="sb-badge">➕</div>
      <div><div class="sb-name">AcciSense</div><div class="sb-sub">Hospital Portal</div></div>
    </div>
  </div>
  <div class="sb-section">Main</div>
  <div class="nav active" onclick="showSection('dashboard',this)"><span class="nav-ic">🏠</span> Overview</div>
  <div class="nav" onclick="showSection('incoming',this)"><span class="nav-ic">🚑</span> Incoming Case<span class="nav-badge" id="live-badge" style="display:none">1</span></div>
  <div class="nav" onclick="showSection('analytics',this)"><span class="nav-ic">📈</span> Analytics</div>
  <div class="nav" onclick="showSection('ranking',this)"><span class="nav-ic">🏆</span> Hospital Rankings</div>
  <div class="sb-section">Hospital</div>
  <div class="nav" onclick="showSection('beds',this)"><span class="nav-ic">🛏</span> Bed Status</div>
  <div class="nav" onclick="showSection('fleet',this)"><span class="nav-ic">🚑</span> Fleet Status</div>
  <div class="nav" onclick="showSection('staff',this)"><span class="nav-ic">👨‍⚕️</span> Staff On Duty</div>
  <div class="nav" onclick="showSection('icu',this)"><span class="nav-ic">🩺</span> ICU Monitor</div>
  <div class="sb-bottom">
    <div class="sb-user">
      <div class="av">🏥</div>
      <div><div class="un">AcciSense</div><div class="ur">Unit ID: HEU-001</div></div>
    </div>
  </div>
</div>

<div class="topbar">
  <div class="tb-title">Hospital Emergency Center</div>
  <div class="tb-right">
    <div class="live-pill"><div class="lp-dot"></div>LIVE</div>
    <div class="tb-time" id="clock">--:--:--</div>
    <div class="bell-wrap" id="bellWrap" onclick="_toggleNotif(event)">
      <button class="bell-btn" title="Notifications">🔔</button>
      <div class="bell-badge hidden" id="bellBadge">0</div>
    </div>
    <a href="/" class="tb-logout">Logout</a>
  </div>
</div>

<div class="main">
  <div id="section-dashboard">
    <div class="page-title">Welcome back, Emergency Unit</div>
    <div class="page-sub">Monitor incoming accident cases and manage hospital response</div>
    <div class="persist-info"><div class="persist-left"><div class="persist-dot"></div><span id="persist-label">Loading...</span></div><span>📅</span></div>
    <div class="stats-row">
      <div class="stat-card"><div class="stat-top"><div class="stat-label">Accepted (All Time)</div><div class="stat-icon-box" style="background:#eff6ff;color:#2563eb;font-size:20px;">✔</div></div><div class="stat-value" id="st-accepted" style="color:#2563eb">0</div><div class="stat-sub">Across all sessions</div></div>
      <div class="stat-card"><div class="stat-top"><div class="stat-label">Declined (All Time)</div><div class="stat-icon-box" style="background:#fef2f2;color:#ef4444;font-size:20px;">✖</div></div><div class="stat-value" id="st-declined" style="color:#ef4444">0</div><div class="stat-sub">Passed to next hospital</div></div>
      <div class="stat-card"><div class="stat-top"><div class="stat-label">Avg Response</div><div class="stat-icon-box" style="background:#fffbeb;color:#d97706">⏱</div></div><div class="stat-value" id="st-avgtime" style="color:#d97706">--</div><div class="stat-sub">Seconds</div></div>
      <div class="stat-card"><div class="stat-top"><div class="stat-label">Active Incidents</div><div class="stat-icon-box" style="background:#fef2f2;color:#ef4444">🚨</div></div><div class="stat-value" id="st-active" style="color:#ef4444">0</div><div class="stat-sub">Requiring action</div></div>
    </div>
    <div id="dash-alert-hosp"></div>
    <div class="charts-row">
      <div class="chart-card"><div class="chart-title">📊 Accept vs Decline</div><canvas id="acceptDoughnut"></canvas></div>
      <div class="chart-card"><div class="chart-title">📉 Hospital Decline Rate (%)</div><canvas id="declineBar"></canvas></div>
    </div>
  </div>

  <div id="section-incoming" style="display:none">
    <div class="page-title">Incoming Case</div><div class="page-sub">Review and respond to incoming accident cases</div>
    <div id="incoming-detail"></div>
  </div>

  <div id="section-analytics" style="display:none">
    <div class="page-title">Hospital Analytics</div><div class="page-sub">Performance metrics across all sessions</div>
    <div class="stats-row">
      <div class="stat-card"><div class="stat-top"><div class="stat-label">Total Requests</div><div class="stat-icon-box" style="background:#eff6ff;color:#3b82f6">📋</div></div><div class="stat-value" id="an-total">0</div><div class="stat-sub">All hospitals</div></div>
      <div class="stat-card"><div class="stat-top"><div class="stat-label">Overall Accept Rate</div><div class="stat-icon-box" style="background:#eff6ff;color:#2563eb;font-size:20px;">✔</div></div><div class="stat-value" id="an-rate" style="color:#2563eb">--%</div><div class="stat-sub">Acceptance %</div></div>
      <div class="stat-card"><div class="stat-top"><div class="stat-label">Fastest Response</div><div class="stat-icon-box" style="background:#eff6ff;color:#3b82f6">⚡</div></div><div class="stat-value" id="an-fast" style="color:#3b82f6">--s</div><div class="stat-sub">Best time recorded</div></div>
      <div class="stat-card"><div class="stat-top"><div class="stat-label">Hospitals Active</div><div class="stat-icon-box" style="background:#faf5ff;color:#8b5cf6">🏥</div></div><div class="stat-value" id="an-active">5</div><div class="stat-sub">In network</div></div>
    </div>
    <div class="charts-row">
      <div class="chart-card"><div class="chart-title">📊 Accepted vs Declined per Hospital</div><canvas id="an-acceptbar"></canvas></div>
      <div class="chart-card"><div class="chart-title">⏱ Avg Response Time (seconds)</div><canvas id="an-rtbar"></canvas></div>
    </div>
    <div class="charts-row">
      <div class="chart-card"><div class="chart-title">🏥 Hospital Accept Rate (%)</div><canvas id="an-rateDonut"></canvas></div>
      <div class="chart-card"><div class="chart-title">📉 Decline % by Hospital</div><canvas id="an-declinebar"></canvas></div>
    </div>
    <div class="table-card">
      <div style="font-size:12px;font-weight:800;margin-bottom:12px;">📋 Hospital Performance Table</div>
      <table><thead><tr><th>Hospital</th><th>Total</th><th>Accepted</th><th>Declined</th><th>Avg Response</th><th>Decline Rate</th><th>Rating</th></tr></thead><tbody id="an-tbody"><tr><td colspan="7" style="text-align:center;color:var(--muted);padding:28px;">No data yet</td></tr></tbody></table>
    </div>
  </div>

  <div id="section-ranking" style="display:none">
    <div class="page-title">Hospital Rankings</div><div class="page-sub">Performance-based leaderboard</div>
    <div class="chart-card"><div style="font-size:12px;font-weight:800;margin-bottom:12px;">🏆 Hospital Performance Rankings</div><div id="ranking-list"></div></div>
  </div>

  <div id="section-beds" style="display:none">
    <div class="page-title">Bed Status</div><div class="page-sub">Real-time bed availability across all wards</div>
    <div class="stats-row">
      <div class="stat-card"><div class="stat-top"><div class="stat-label">Total Beds</div><div class="stat-icon-box" style="background:#eff6ff;color:#3b82f6">🛏</div></div><div class="stat-value" id="bed-total">0</div></div>
      <div class="stat-card"><div class="stat-top"><div class="stat-label">Occupied</div><div class="stat-icon-box" style="background:#fef2f2;color:#ef4444">🔴</div></div><div class="stat-value" id="bed-occ" style="color:#ef4444">0</div></div>
      <div class="stat-card"><div class="stat-top"><div class="stat-label">Available</div><div class="stat-icon-box" style="background:#eff6ff;color:#2563eb">🟢</div></div><div class="stat-value" id="bed-avail" style="color:#2563eb">0</div></div>
      <div class="stat-card"><div class="stat-top"><div class="stat-label">Occupancy Rate</div><div class="stat-icon-box" style="background:#fffbeb;color:#d97706">📊</div></div><div class="stat-value" id="bed-rate" style="color:#d97706">0%</div></div>
    </div>
    <div class="bed-grid" id="bed-grid"></div>
    <div class="charts-row">
      <div class="chart-card"><div class="chart-title">🛏 Bed Occupancy by Ward</div><canvas id="bedBar"></canvas></div>
      <div class="chart-card"><div class="chart-title">📊 Available vs Occupied</div><canvas id="bedDoughnut"></canvas></div>
    </div>
  </div>

  <div id="section-fleet" style="display:none">
    <div class="page-title">Fleet Status</div><div class="page-sub">Emergency vehicle tracking</div>
    <div class="stats-row">
      <div class="stat-card"><div class="stat-top"><div class="stat-label">Total Fleet</div><div class="stat-icon-box" style="background:#eff6ff;color:#3b82f6">🚑</div></div><div class="stat-value" id="fl-total">0</div></div>
      <div class="stat-card"><div class="stat-top"><div class="stat-label">Available</div><div class="stat-icon-box" style="background:#eff6ff;color:#2563eb;font-size:20px;">✔</div></div><div class="stat-value" id="fl-avail" style="color:#2563eb">0</div></div>
      <div class="stat-card"><div class="stat-top"><div class="stat-label">On Mission</div><div class="stat-icon-box" style="background:#eff6ff;color:#3b82f6">🚑</div></div><div class="stat-value" id="fl-disp" style="color:#3b82f6">0</div></div>
      <div class="stat-card"><div class="stat-top"><div class="stat-label">Maintenance</div><div class="stat-icon-box" style="background:#fffbeb;color:#d97706">🔧</div></div><div class="stat-value" id="fl-maint" style="color:#d97706">0</div></div>
    </div>
    <div class="table-card">
      <div style="font-size:12px;font-weight:800;margin-bottom:12px;">🚑 Ambulance Fleet</div>
      <table><thead><tr><th>Vehicle ID</th><th>Type</th><th>Driver</th><th>Location</th><th>Last Trip</th><th>Status</th></tr></thead><tbody id="fl-tbody"></tbody></table>
    </div>
  </div>

  <div id="section-staff" style="display:none">
    <div class="page-title">Staff On Duty</div><div class="page-sub">Current medical staff on emergency duty</div>
    <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:14px;">
      <div class="chart-card" style="display:flex;align-items:center;gap:12px;"><div style="font-size:30px;">👨‍⚕️</div><div><div style="font-weight:800;font-size:13px;">Dr. Arjun Mehta</div><div style="font-size:11px;color:var(--muted);">Emergency Surgeon</div><span class="badge badge-success" style="margin-top:5px;display:inline-flex;">● On Duty</span></div></div>
      <div class="chart-card" style="display:flex;align-items:center;gap:12px;"><div style="font-size:30px;">👩‍⚕️</div><div><div style="font-weight:800;font-size:13px;">Dr. Priya Sharma</div><div style="font-size:11px;color:var(--muted);">Trauma Specialist</div><span class="badge badge-success" style="margin-top:5px;display:inline-flex;">● On Duty</span></div></div>
      <div class="chart-card" style="display:flex;align-items:center;gap:12px;"><div style="font-size:30px;">👨‍⚕️</div><div><div style="font-weight:800;font-size:13px;">Dr. Ravi Kumar</div><div style="font-size:11px;color:var(--muted);">Cardiologist</div><span class="badge badge-warn" style="margin-top:5px;display:inline-flex;">⚠ Break</span></div></div>
      <div class="chart-card" style="display:flex;align-items:center;gap:12px;"><div style="font-size:30px;">👩‍⚕️</div><div><div style="font-weight:800;font-size:13px;">Dr. Sunita Patel</div><div style="font-size:11px;color:var(--muted);">Neurologist</div><span class="badge badge-success" style="margin-top:5px;display:inline-flex;">● On Duty</span></div></div>
      <div class="chart-card" style="display:flex;align-items:center;gap:12px;"><div style="font-size:30px;">👨‍⚕️</div><div><div style="font-weight:800;font-size:13px;">Dr. Anil Reddy</div><div style="font-size:11px;color:var(--muted);">Orthopedic Surgeon</div><span class="badge badge-success" style="margin-top:5px;display:inline-flex;">● On Duty</span></div></div>
      <div class="chart-card" style="display:flex;align-items:center;gap:12px;"><div style="font-size:30px;">👩‍⚕️</div><div><div style="font-weight:800;font-size:13px;">Dr. Kavya Nair</div><div style="font-size:11px;color:var(--muted);">ICU Specialist</div><span class="badge badge-danger" style="margin-top:5px;display:inline-flex;">● Off Duty</span></div></div>
    </div>
  </div>

  <div id="section-icu" style="display:none">
    <div class="page-title">ICU Monitor</div><div class="page-sub">Intensive care unit patient monitoring</div>
    <div class="stats-row">
      <div class="stat-card"><div class="stat-top"><div class="stat-label">ICU Beds Total</div><div class="stat-icon-box" style="background:#faf5ff;color:#8b5cf6">🩺</div></div><div class="stat-value">20</div></div>
      <div class="stat-card"><div class="stat-top"><div class="stat-label">ICU Occupied</div><div class="stat-icon-box" style="background:#fef2f2;color:#ef4444">🔴</div></div><div class="stat-value" style="color:#ef4444">14</div></div>
      <div class="stat-card"><div class="stat-top"><div class="stat-label">ICU Available</div><div class="stat-icon-box" style="background:#eff6ff;color:#2563eb">🟢</div></div><div class="stat-value" style="color:#2563eb">6</div></div>
      <div class="stat-card"><div class="stat-top"><div class="stat-label">Critical Patients</div><div class="stat-icon-box" style="background:#fef2f2;color:#ef4444">⚠️</div></div><div class="stat-value" style="color:#ef4444">3</div></div>
    </div>
    <div class="table-card">
      <div style="font-size:12px;font-weight:800;margin-bottom:12px;">🩺 ICU Patient Status</div>
      <table><thead><tr><th>Bed #</th><th>Patient ID</th><th>Condition</th><th>Admitted</th><th>Doctor</th><th>Status</th></tr></thead>
      <tbody>
        <tr><td>ICU-01</td><td>P-2401</td><td>Head Trauma</td><td>09:15 AM</td><td>Dr. Arjun Mehta</td><td><span class="badge badge-danger">Critical</span></td></tr>
        <tr><td>ICU-02</td><td>P-2402</td><td>Cardiac Arrest</td><td>10:30 AM</td><td>Dr. Ravi Kumar</td><td><span class="badge badge-danger">Critical</span></td></tr>
        <tr><td>ICU-03</td><td>P-2403</td><td>Spinal Injury</td><td>08:45 AM</td><td>Dr. Anil Reddy</td><td><span class="badge badge-warn">Serious</span></td></tr>
        <tr><td>ICU-04</td><td>P-2404</td><td>Internal Bleeding</td><td>11:00 AM</td><td>Dr. Priya Sharma</td><td><span class="badge badge-warn">Serious</span></td></tr>
        <tr><td>ICU-05</td><td>P-2405</td><td>Fractures</td><td>07:20 AM</td><td>Dr. Anil Reddy</td><td><span class="badge badge-blue">Stable</span></td></tr>
      </tbody></table>
    </div>
  </div>
</div>

<script>
""" + BELL_JS + """
var CS={},lastCases=0;
var acceptDChart,declineBarC,anAccBar,anRtBar,anRateD,anDecBar,bedBarC,bedDC;
var bedChartInit=false,anChartInit=false;

setInterval(function(){document.getElementById('clock').textContent=new Date().toLocaleTimeString('en-IN',{hour12:false});},1000);

function showSection(name,el){
  document.querySelectorAll('[id^="section-"]').forEach(function(s){s.style.display='none';});
  document.getElementById('section-'+name).style.display='block';
  document.querySelectorAll('.nav').forEach(function(n){n.classList.remove('active');});
  if(el)el.classList.add('active');
  if(name==='analytics')renderAnalytics();
  if(name==='ranking')renderRanking();
  if(name==='beds')renderBeds();
  if(name==='fleet')renderFleet();
  _closeNotif();
}

function showToast(msg){var t=document.getElementById('toast');document.getElementById('toast-msg').textContent=msg;t.classList.add('show');setTimeout(function(){t.classList.remove('show');},6000);}
function fmtTime(iso){return iso?new Date(iso).toLocaleTimeString('en-IN',{hour12:false}):'--';}
function getVI(p){var VT={TN01:{type:"Sedan",icon:"🚗"},TN02:{type:"SUV",icon:"🚙"},TN03:{type:"Truck",icon:"🚛"},MH:{type:"Motorcycle",icon:"🏍️"},KA:{type:"Auto-Rickshaw",icon:"🛺"},AP:{type:"Bus",icon:"🚌"},TS:{type:"Van",icon:"🚐"},DL:{type:"Hatchback",icon:"🚘"},GJ:{type:"Pickup",icon:"🛻"},WB:{type:"Taxi",icon:"🚕"}};for(var k in VT){if(p.toUpperCase().startsWith(k))return VT[k];}return{type:"Vehicle",icon:"🚗"};}

function buildHospCard(s){
  if(!s.current_case)return'<div class="no-case"><div class="no-case-icon">🏥</div><div class="no-case-text">No incoming case — System ready</div></div>';
  var c=s.current_case,idx=s.hospital_index;
  var HOSPS=[{name:"Apollo Emergency Center",dist:"1.2 km"},{name:"City Care Hospital",dist:"2.4 km"},{name:"Metro Trauma Hospital",dist:"3.1 km"},{name:"Green Cross Medical",dist:"4.8 km"},{name:"National Emergency Hospital",dist:"6.2 km"}];
  var thisH=idx<HOSPS.length?HOSPS[idx]:null;
  var showBtns=s.case_status==='pending'&&thisH;
  var isPending=s.case_status==='pending',isDispatched=s.case_status==='ambulance_dispatched';
  var v1=getVI(c.vehicle_1),v2=getVI(c.vehicle_2);
  var pill=isPending?'<div class="disp-pill-emergency">🚨 INCOMING CASE</div>':'<div class="disp-pill-resolved">✅ RESOLVED</div>';
  var badge=isDispatched?'<div class="disp-status-badge dsb-green">AMBULANCE DISPATCHED</div>':isPending?'<div class="disp-status-badge dsb-red">PENDING</div>':'<div class="disp-status-badge dsb-orange">NO HOSPITAL</div>';
  var at=fmtTime(s.alert_time),act=fmtTime(s.accept_time);
  var hero=isDispatched?'<div class="wave-hero blue-wave"><div class="wave-vehicle">🚑</div><div class="wave-info"><h3>Ambulance Dispatched</h3><p>Alert: '+at+' &nbsp;|&nbsp; Accepted: '+act+'</p></div></div>':isPending?'<div class="wave-hero amber-wave"><div class="wave-vehicle">🚨</div><div class="wave-info"><h3 style="color:#78350f;">Emergency Alert</h3><p style="color:#d97706;">Received at '+at+' — Awaiting acceptance</p></div></div>':'';
  var btns=showBtns?'<div class="action-btns"><form action="/accept" method="post" style="flex:1"><button class="btn-accept" type="submit">✅ Accept & Dispatch Ambulance</button></form><form action="/decline" method="post" style="flex:1"><button class="btn-decline" type="submit">❌ Decline — Next Hospital</button></form></div>':'';
  return'<div class="disp-card-wrap '+(isPending?'emergency':'resolved')+'">'
    +'<div class="disp-header">'+pill+'<div class="disp-title">Emergency Accident Response Required</div>'+badge+'</div>'
    +hero
    +'<div class="info-grid-icon">'
    +'<div class="igf"><div class="igf-icon" style="background:#eff6ff">🏥</div><div><div class="igf-label">Hospital Contacted</div><div class="igf-value">'+(thisH?thisH.name:(HOSPS[Math.max(0,idx-1)]||{name:'N/A'}).name)+'</div></div></div>'
    +'<div class="igf"><div class="igf-icon" style="background:#eff6ff">📍</div><div><div class="igf-label">Distance</div><div class="igf-value">'+(thisH?thisH.dist:'N/A')+'</div></div></div>'
    +'<div class="igf"><div class="igf-icon" style="background:#fffbeb">⏰</div><div><div class="igf-label">Alert Time</div><div class="igf-value">'+at+'</div></div></div>'
    +'<div class="igf"><div class="igf-icon" style="background:#eff6ff">🚗</div><div><div class="igf-label">Vehicles</div><div class="igf-value">'+v1.icon+' '+c.vehicle_1+' / '+v2.icon+' '+c.vehicle_2+'</div></div></div>'
    +'<div class="igf"><div class="igf-icon" style="background:#eff6ff">🌐</div><div><div class="igf-label">GPS Location</div><div class="igf-value">'+(c.latitude||'N/A')+', '+(c.longitude||'N/A')+'</div></div></div>'
    +'<div class="igf"><div class="igf-icon" style="background:#faf5ff">🚓</div><div><div class="igf-label">Police Assigned</div><div class="igf-value">'+(s.selected_police||'N/A')+'</div></div></div>'
    +'</div>'+btns+'</div>';
}

function initDashCharts(){
  Chart.defaults.font.family='Nunito';Chart.defaults.font.weight='700';Chart.defaults.color='#64748b';
  acceptDChart=new Chart(document.getElementById('acceptDoughnut'),{type:'doughnut',data:{labels:['Accepted','Declined'],datasets:[{data:[0,0],backgroundColor:['rgba(37,99,235,0.75)','rgba(239,68,68,0.7)'],borderWidth:0}]},options:{plugins:{legend:{position:'bottom',labels:{boxWidth:10,padding:8}}},cutout:'60%',animation:false}});
  declineBarC=new Chart(document.getElementById('declineBar'),{type:'bar',data:{labels:['Apollo','City Care','Metro','Green Cross','National'],datasets:[{label:'Decline %',data:[0,0,0,0,0],backgroundColor:'rgba(239,68,68,0.5)',borderColor:'#ef4444',borderWidth:1,borderRadius:6}]},options:{plugins:{legend:{display:false}},indexAxis:'y',scales:{x:{beginAtZero:true,max:100},y:{grid:{display:false}}},animation:false}});
}

async function renderAnalytics(){
  var r=await fetch('/api/hospital-stats');var stats=await r.json();
  var names=stats.map(function(s){return s.name.split(' ')[0];});
  var acc=stats.map(function(s){return s.accepted;}),dec=stats.map(function(s){return s.declined;}),avgT=stats.map(function(s){return s.avg_time;}),decP=stats.map(function(s){return s.decline_pct;});
  var totR=stats.reduce(function(a,s){return a+s.total;},0),totA=stats.reduce(function(a,s){return a+s.accepted;},0);
  var rate=totR>0?Math.round(totA/totR*100):0;
  var times=stats.filter(function(s){return s.avg_time>0;}).map(function(s){return s.avg_time;});
  var fast=times.length?Math.min.apply(null,times):null;
  document.getElementById('an-total').textContent=totR;document.getElementById('an-rate').textContent=rate+'%';document.getElementById('an-fast').textContent=fast?fast+'s':'--s';
  if(!anChartInit){
    anChartInit=true;
    anAccBar=new Chart(document.getElementById('an-acceptbar'),{type:'bar',data:{labels:names,datasets:[{label:'Accepted',data:acc,backgroundColor:'rgba(37,99,235,0.65)',borderColor:'#2563eb',borderWidth:1,borderRadius:6},{label:'Declined',data:dec,backgroundColor:'rgba(239,68,68,0.6)',borderColor:'#ef4444',borderWidth:1,borderRadius:6}]},options:{plugins:{legend:{labels:{boxWidth:10,font:{size:11}}}},scales:{y:{beginAtZero:true},x:{grid:{display:false}}},animation:false}});
    anRtBar=new Chart(document.getElementById('an-rtbar'),{type:'bar',data:{labels:names,datasets:[{label:'Avg Response(s)',data:avgT,backgroundColor:'rgba(59,130,246,0.6)',borderColor:'#3b82f6',borderWidth:1,borderRadius:6}]},options:{plugins:{legend:{display:false}},scales:{y:{beginAtZero:true},x:{grid:{display:false}}},animation:false}});
    anRateD=new Chart(document.getElementById('an-rateDonut'),{type:'doughnut',data:{labels:names,datasets:[{data:stats.map(function(s){return s.total>0?Math.round(s.accepted/s.total*100):0;}),backgroundColor:['#2563eb','#3b82f6','#60a5fa','#93c5fd','#bfdbfe'],borderWidth:0}]},options:{plugins:{legend:{position:'bottom',labels:{boxWidth:10,font:{size:11}}}},cutout:'55%',animation:false}});
    anDecBar=new Chart(document.getElementById('an-declinebar'),{type:'bar',data:{labels:names,datasets:[{label:'Decline %',data:decP,backgroundColor:'rgba(239,68,68,0.5)',borderColor:'#ef4444',borderWidth:1,borderRadius:6}]},options:{plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,max:100},x:{grid:{display:false}}},animation:false}});
  }else{anAccBar.data.datasets[0].data=acc;anAccBar.data.datasets[1].data=dec;anAccBar.update();anRtBar.data.datasets[0].data=avgT;anRtBar.update();anDecBar.data.datasets[0].data=decP;anDecBar.update();}
  var tbody=document.getElementById('an-tbody');
  tbody.innerHTML=stats.map(function(s,i){var bc=s.decline_pct>50?'badge-danger':s.decline_pct>20?'badge-warn':'badge-success';var stars='★'.repeat(Math.max(1,5-i));return'<tr><td style="font-weight:800">'+s.name+'</td><td>'+s.total+'</td><td style="color:#2563eb;font-weight:800">'+s.accepted+'</td><td style="color:#ef4444;font-weight:800">'+s.declined+'</td><td>'+(s.avg_time>0?s.avg_time+'s':'--')+'</td><td><span class="badge '+bc+'">'+s.decline_pct+'%</span></td><td style="color:#f59e0b">'+stars+'</td></tr>';}).join('');
}

async function renderRanking(){
  var r=await fetch('/api/hospital-stats');var stats=await r.json();
  var sorted=stats.slice().sort(function(a,b){if(a.avg_time===0&&b.avg_time===0)return b.accepted-a.accepted;if(a.avg_time===0)return 1;if(b.avg_time===0)return -1;return a.avg_time-b.avg_time;});
  var maxD=Math.max.apply(null,stats.map(function(s){return s.decline_pct;}).concat([0]));
  var medals=['🥇','🥈','🥉'];
  document.getElementById('ranking-list').innerHTML=sorted.map(function(h,i){return'<div class="ranking-row"><div class="rank-num">'+(medals[i]||i+1)+'</div><div class="rank-name"><div class="rank-name-main">'+h.name+'</div><div class="rank-name-sub">Accepted: '+h.accepted+' | Declined: '+h.declined+' | Avg: '+(h.avg_time>0?h.avg_time+'s':'--')+'</div></div><div style="display:flex;gap:5px;align-items:center">'+(i===0&&h.avg_time>0?'<span class="hosp-tag tag-fast">⚡ Fastest</span>':'')+(h.decline_pct===maxD&&maxD>0?'<span class="hosp-tag tag-slow">⚠ Most Declined</span>':'')+'</div><div class="hosp-stars">'+'★'.repeat(Math.max(1,5-i))+'</div><div style="font-size:12px;color:'+(h.decline_pct>50?'#ef4444':'#2563eb')+';min-width:46px;text-align:right;font-weight:800">'+h.decline_pct+'%</div></div>';}).join('');
}

async function renderBeds(){
  var r=await fetch('/api/bed-status');var beds=await r.json();
  var tot=0,occ=0,avail=0;beds.forEach(function(b){tot+=b.total;occ+=b.occupied;avail+=b.available;});
  document.getElementById('bed-total').textContent=tot;document.getElementById('bed-occ').textContent=occ;document.getElementById('bed-avail').textContent=avail;document.getElementById('bed-rate').textContent=Math.round(occ/tot*100)+'%';
  document.getElementById('bed-grid').innerHTML=beds.map(function(b){var pct=Math.round(b.occupied/b.total*100);return'<div class="bed-card '+b.status+'"><div class="bed-ward">'+b.ward+'</div><div class="bed-numbers"><div class="bed-avail '+b.status+'">'+b.available+'</div><div class="bed-total">/ '+b.total+'</div></div><div class="bed-bar"><div class="bed-fill '+b.status+'" style="width:'+pct+'%"></div></div><div class="bed-label">'+pct+'% occupied</div></div>';}).join('');
  if(!bedChartInit){
    bedChartInit=true;
    bedBarC=new Chart(document.getElementById('bedBar'),{type:'bar',data:{labels:beds.map(function(b){return b.ward.replace(' Ward','').replace(' Unit','');}),datasets:[{label:'Occupied',data:beds.map(function(b){return b.occupied;}),backgroundColor:'rgba(239,68,68,0.6)',borderColor:'#ef4444',borderWidth:1,borderRadius:4},{label:'Available',data:beds.map(function(b){return b.available;}),backgroundColor:'rgba(37,99,235,0.6)',borderColor:'#2563eb',borderWidth:1,borderRadius:4}]},options:{plugins:{legend:{labels:{boxWidth:10,font:{size:11}}}},scales:{y:{beginAtZero:true},x:{grid:{display:false},ticks:{maxRotation:45,font:{size:9}}}},animation:false}});
    bedDC=new Chart(document.getElementById('bedDoughnut'),{type:'doughnut',data:{labels:['Occupied','Available'],datasets:[{data:[occ,avail],backgroundColor:['rgba(239,68,68,0.7)','rgba(37,99,235,0.65)'],borderWidth:0}]},options:{plugins:{legend:{position:'bottom',labels:{boxWidth:10,font:{size:12}}}},cutout:'60%',animation:false}});
  }
}

async function renderFleet(){
  var r=await fetch('/api/fleet-status');var data=await r.json();
  var av=0,di=0,ma=0;data.forEach(function(f){if(f.status==='available')av++;else if(f.status==='dispatched')di++;else ma++;});
  document.getElementById('fl-total').textContent=data.length;document.getElementById('fl-avail').textContent=av;document.getElementById('fl-disp').textContent=di;document.getElementById('fl-maint').textContent=ma;
  document.getElementById('fl-tbody').innerHTML=data.map(function(f){var bc=f.status==='available'?'badge-success':f.status==='dispatched'?'badge-blue':'badge-warn';var lbl=f.status==='available'?'✅ Available':f.status==='dispatched'?'🚑 Dispatched':'🔧 Maintenance';return'<tr><td><strong>'+f.id+'</strong></td><td>'+f.type+'</td><td>'+f.driver+'</td><td>'+f.location+'</td><td>'+f.last_trip+'</td><td><span class="badge '+bc+'">'+lbl+'</span></td></tr>';}).join('');
}

async function poll(){
  try{
    var r=await fetch('/api/state');var s=await r.json();CS=s;
    var hs=s.hospital_stats||{},totA=0,totD=0,totT=0,totN=0;
    Object.values(hs).forEach(function(h){totA+=h.accepted;totD+=h.declined;h.times.forEach(function(t){totT+=t;totN++;});});
    document.getElementById('st-accepted').textContent=totA;
    document.getElementById('st-declined').textContent=totD;
    document.getElementById('st-avgtime').textContent=totN>0?Math.round(totT/totN):'--';
    document.getElementById('st-active').textContent=s.case_status==='pending'?1:0;
    var hist=s.case_history||[];
    var earliest=hist.length?hist[0].date:'—',latest=hist.length?hist[hist.length-1].date:'—';
    document.getElementById('persist-label').textContent=hist.length?'Persistent data loaded — '+hist.length+' total cases from '+earliest+' to '+latest:'Persistent storage active — No cases yet recorded';
    var lb=document.getElementById('live-badge');lb.style.display=s.case_status==='pending'?'inline':'none';
    if(s.current_case&&s.total_cases>lastCases&&s.case_status==='pending'){
      showToast('🏥 Emergency case incoming — immediate response required');
      _pushNotif('🚨','#fef2f2','New Emergency Case!','Accident involving '+s.current_case.vehicle_1+' & '+s.current_case.vehicle_2+' requires response.');
    }
    if(s.total_cases>0)lastCases=s.total_cases;
    var hospHtml=buildHospCard(s);
    document.getElementById('dash-alert-hosp').innerHTML=hospHtml;
    document.getElementById('incoming-detail').innerHTML=hospHtml;
    if(acceptDChart){acceptDChart.data.datasets[0].data=[totA,totD];acceptDChart.update();}
    if(declineBarC){
      var sr=await fetch('/api/hospital-stats');var sd=await sr.json();
      declineBarC.data.datasets[0].data=sd.map(function(st){return st.decline_pct;});declineBarC.update();
    }
  }catch(e){console.error(e);}
}

// Init bell with blue-theme hospital notifications
_initBell('#2563eb','#eff6ff',[
  {icon:'🚨',bg:'#fef2f2',title:'Emergency Alert Active',desc:'Incoming accident case awaiting hospital response.',time:'Just now',unread:true},
  {icon:'🛏',bg:'#fffbeb',title:'Cardiac ICU Critical',desc:'Cardiac ICU at 92% capacity — only 1 bed remaining.',time:'8 min ago',unread:true},
  {icon:'🚑',bg:'#eff6ff',title:'AMB-002 Dispatched',desc:'Ambulance en route to accident site on NH44.',time:'15 min ago',unread:false},
  {icon:'✅',bg:'#eff6ff',title:'Previous Case Resolved',desc:'Last emergency case successfully dispatched.',time:'32 min ago',unread:false},
  {icon:'📋',bg:'#eff6ff',title:'System Online',desc:'All hospital systems operational and connected.',time:'1 hr ago',unread:false},
]);
initDashCharts();
poll();
setInterval(poll,2000);
</script>
</body>
</html>"""