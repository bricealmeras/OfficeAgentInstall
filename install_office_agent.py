import os
import sys

# Chemins Antigravity
home = os.path.expanduser("~")
workspace = os.path.join(home, ".gemini", "antigravity", "scratch", "office-agent-ipc")
skills_dir = os.path.join(home, ".gemini", "antigravity", "config", "skills", "office-agent")

os.makedirs(os.path.join(workspace, "responses"), exist_ok=True)
os.makedirs(skills_dir, exist_ok=True)

# 1. gateway_ipc.py
GATEWAY_CODE = """iimport os, sys, json, time, uuid, asyncio
from typing import Dict, Any
from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, StreamingResponse
import uvicorn

app = FastAPI(title="Office Agent IPC Gateway")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_credentials=True, allow_methods=["*"], allow_headers=["*"])

WORKSPACE_DIR = os.path.join(os.path.expanduser("~"), ".gemini", "antigravity", "scratch", "office-agent-ipc")
REQUESTS_FILE = os.path.join(WORKSPACE_DIR, "requests.jsonl")
RESPONSES_DIR = os.path.join(WORKSPACE_DIR, "responses")

os.makedirs(RESPONSES_DIR, exist_ok=True)
# Purge des fichiers de session précédente au démarrage
open(REQUESTS_FILE, 'w').close()
STATE_FILE = os.path.join(WORKSPACE_DIR, "last_line.txt")
if os.path.exists(STATE_FILE):
    os.remove(STATE_FILE)
# Purge des réponses orphelines
for f in os.listdir(RESPONSES_DIR):
    try: os.remove(os.path.join(RESPONSES_DIR, f))
    except: pass

def extract_prompt(payload: Dict[str, Any]) -> str:
    messages = payload.get("messages", [])
    last_user_text = ""
    office_context = ""

    # Construire l'historique des échanges (max 6 derniers messages, hors le courant)
    history_lines = []
    SKIP_INTERNAL = {"context_snip", "retrieve_snipped", "read_surface_context", "update_setting", "todo_write"}
    for msg in messages[:-1]:  # tout sauf le dernier message (le courant)
        role = msg.get("role")
        content = msg.get("content")
        if isinstance(content, str) and content.strip():
            label = "Utilisateur" if role == "user" else "Assistant"
            history_lines.append(f"{label}: {content.strip()[:300]}")
        elif isinstance(content, list):
            for c in content:
                if c.get("type") == "text" and c.get("text","").strip():
                    text = c["text"].split("[id:")[0].strip()[:300]
                    if text:
                        label = "Utilisateur" if role == "user" else "Assistant"
                        history_lines.append(f"{label}: {text}")
                        break
    # Garder les 6 derniers échanges max
    history_lines = history_lines[-6:]

    # Extraire le dernier message utilisateur et le contexte Office
    for msg in reversed(messages):
        if msg.get("role") == "user":
            content = msg.get("content")
            if isinstance(content, list):
                for c in content:
                    if c.get("type") == "text" and not last_user_text:
                        last_user_text = c.get("text", "").split("[id:")[0].strip()
                    elif c.get("type") == "tool_result" and not office_context:
                        office_context = c.get("content", "")
            if last_user_text: break

    prompt = last_user_text or "Bonjour"

    # Injecter l'historique s'il existe
    if history_lines:
        prompt = "[Historique de la conversation]\n" + "\n".join(history_lines) + "\n\n[Message actuel]\n" + prompt

    if office_context:
        prompt += f"\n\n[Contexte Office] :\n{office_context}"

    # Injecter la liste des outils disponibles !
    tools = payload.get("tools", [])
    if tools:
        prompt += "\n\n[Outils Office Disponibles] :\nTu peux utiliser les outils suivants si besoin en retournant EXACTEMENT un JSON du type {\"tool\": \"nom_outil\", \"param1\": \"valeur\"}. N'utilise que les paramètres requis.\n"
        for t in tools:
            name = t.get("name")
            if name in ["context_snip", "retrieve_snipped", "read_surface_context", "update_setting", "todo_write"]:
                continue # Ignorer les outils internes d'interface
            schema = t.get("input_schema", {}).get("properties", {})
            params = [f'"{k}": ...' for k in schema.keys() if k != "explanation"]

            desc = t.get('description', '')[:150]
            if name == "execute_office_js":
                desc = "Permet d'exécuter du code JavaScript/TypeScript (Office.js) ! Le code reçoit `context: Excel.RequestContext`. N'oublie pas `await context.sync();` à la fin !"

            prompt += f"- {name} (Params: {', '.join(params)}) : {desc}\n"
            
    return prompt

@app.get("/v1")
@app.get("/v1/models")
@app.get("/models")
@app.get("/v1/v1/models")
async def mock_models():
    return {"object": "list", "data": [{"id": "claude-3-5-sonnet-20241022", "object": "model", "created": 1729555200, "owned_by": "antigravity"}]}

@app.post("/v1/messages")
@app.post("/messages")
@app.post("/v1/v1/messages")
async def handle_messages(request: Request):
    payload = await request.json()
    model = payload.get("model", "claude-3-5-sonnet-20241022")
    prompt = extract_prompt(payload)
    req_id = f"req_{uuid.uuid4().hex[:8]}"
    
    with open(REQUESTS_FILE, "a", encoding="utf-8") as f:
        f.write(json.dumps({"id": req_id, "prompt": prompt}) + "\n")
        
    response_file = os.path.join(RESPONSES_DIR, f"{req_id}.json")
    start_time = time.time()
    agent_text = ""
    
    while time.time() - start_time < 120:
        if os.path.exists(response_file):
            try:
                with open(response_file, "r", encoding="utf-8") as f:
                    resp_data = json.load(f)
                    agent_text = resp_data.get("response", json.dumps(resp_data)) if "response" in resp_data else json.dumps(resp_data)
                os.remove(response_file)
                break
            except:
                pass
        await asyncio.sleep(0.1)
        
    if not agent_text:
        agent_text = "Désolé, l'agent n'a pas répondu à temps."

    if payload.get("stream", False):
        async def generate_claude_stream(text: str, model_name: str):
            msg_id = f"msg_{uuid.uuid4().hex[:16]}"
            yield f'event: message_start\ndata: {json.dumps({"type": "message_start", "message": {"id": msg_id, "type": "message", "role": "assistant", "model": model_name, "content": [], "stop_reason": None, "stop_sequence": None, "usage": {"input_tokens": 10, "output_tokens": 1}}})}\n\n'
            
            tool_call = None
            if text.strip().startswith("{") and text.strip().endswith("}"):
                try:
                    parsed = json.loads(text)
                    if "tool" in parsed: tool_call = parsed
                except: pass

            if tool_call:
                tool_name = tool_call.pop("tool")
                if "explanation" not in tool_call: tool_call["explanation"] = "Exécution pour Office"
                yield f'event: content_block_start\ndata: {json.dumps({"type": "content_block_start", "index": 0, "content_block": {"type": "tool_use", "id": f"toolu_{uuid.uuid4().hex[:16]}", "name": tool_name, "input": tool_call}})}\n\n'
                yield f'event: content_block_stop\ndata: {json.dumps({"type": "content_block_stop", "index": 0})}\n\n'
            else:
                yield f'event: content_block_start\ndata: {json.dumps({"type": "content_block_start", "index": 0, "content_block": {"type": "text", "text": ""}})}\n\n'
                for i in range(0, len(text.split(" ")), 5):
                    chunk = " ".join(text.split(" ")[i:i+5]) + " "
                    yield f'event: content_block_delta\ndata: {json.dumps({"type": "content_block_delta", "index": 0, "delta": {"type": "text_delta", "text": chunk}})}\n\n'
                    await asyncio.sleep(0.02)
                yield f'event: content_block_stop\ndata: {json.dumps({"type": "content_block_stop", "index": 0})}\n\n'
            
            yield f'event: message_delta\ndata: {json.dumps({"type": "message_delta", "delta": {"stop_reason": "tool_use" if tool_call else "end_turn", "stop_sequence": None}, "usage": {"output_tokens": 50}})}\n\n'
            yield f'event: message_stop\ndata: {json.dumps({"type": "message_stop"})}\n\n'
            
        return StreamingResponse(generate_claude_stream(agent_text, model), media_type="text/event-stream")
    else:
        return JSONResponse(content={"id": f"msg_{uuid.uuid4().hex[:16]}", "type": "message", "role": "assistant", "model": model, "content": [{"type": "text", "text": agent_text}], "stop_reason": "end_turn", "stop_sequence": None, "usage": {"input_tokens": 10, "output_tokens": len(agent_text.split())}})

if __name__ == "__main__":
    uvicorn.run(app, host="127.0.0.1", port=8000)
"""

# 2. poll.py
POLL_CODE = """import os, time, sys, json

WORKSPACE_DIR = os.path.join(os.path.expanduser("~"), ".gemini", "antigravity", "scratch", "office-agent-ipc")
REQUESTS_FILE = os.path.join(WORKSPACE_DIR, "requests.jsonl")
STATE_FILE = os.path.join(WORKSPACE_DIR, "last_line.txt")

last_line = 0
if os.path.exists(STATE_FILE):
    with open(STATE_FILE, "r") as f:
        last_line = int(f.read().strip() or "0")

while True:
    if os.path.exists(REQUESTS_FILE):
        with open(REQUESTS_FILE, "r", encoding="utf-8") as f:
            lines = f.readlines()
        if len(lines) > last_line:
            new_request_str = lines[last_line].strip()
            new_request = json.loads(new_request_str)
            with open(STATE_FILE, "w") as f:
                f.write(str(last_line + 1))
            with open(os.path.join(WORKSPACE_DIR, 'current_request.json'), 'w', encoding='utf-8') as f:
                json.dump(new_request, f, ensure_ascii=False, indent=2)
            print(f"NOUVELLE REQUETE ID: {new_request['id']}")
            print("Lisez le fichier current_request.json pour voir le prompt et les outils.")
            sys.exit(0)
    time.sleep(0.1)
"""

# 3. SKILL.md
SKILL_CODE = """---
name: office-agent
description: Démarrer l'agent Antigravity en arrière-plan pour la suite Microsoft Office (Word, Excel, PowerPoint).
---
# Lancement de l'Agent Office (IPC)

Lorsque l'utilisateur tape `/office-agent`, tu dois créer et invoquer le sous-agent qui tourne en arrière-plan.

## Étape 1 : Définir le sous-agent
Utilise l'outil `define_subagent` :
- **name** : `office_backend`
- **description** : Office backend processor
- **enable_write_tools** : true
- **enable_mcp_tools** : false
- **system_prompt** : 
```text
Tu es le "Cerveau Office" d'Antigravity.
Ton rôle est d'écouter les requêtes venant de Word, Excel, ou PowerPoint en tâche de fond.

Dès que tu es invoqué POUR LA TOUTE PREMIÈRE FOIS :
1. Lance IMMÉDIATEMENT la passerelle réseau en exécutant :
`python C:\\\\Users\\\\<VOTRE_NOM>\\\\.gemini\\\\antigravity\\\\scratch\\\\office-agent-ipc\\\\gateway_ipc.py`
(utilise run_command avec IsDaemon=true, c'est un serveur persistant).
2. Ensuite, lance ton écouteur avec :
`python C:\\\\Users\\\\<VOTRE_NOM>\\\\.gemini\\\\antigravity\\\\scratch\\\\office-agent-ipc\\\\poll.py`
(utilise run_command avec IsDaemon=false et WaitMsBeforeAsync=1000).

Dès que l'utilisateur pose une question dans Office, le script s'arrête et t'envoie la ligne JSON.
Le prompt contiendra le texte de l'utilisateur, le [Contexte Office], et surtout la liste des [Outils Office Disponibles] pour ce logiciel spécifique !

TRES IMPORTANT (UTILISATION DES OUTILS) : Si tu as besoin d'utiliser un des outils proposés, ta réponse DOIT ÊTRE EXACTEMENT ce format JSON (sans aucun texte autour) :
{"tool": "nom_de_l_outil", "param_requis": "valeur"}
(L'outil sera exécuté par le logiciel et le résultat te sera renvoyé à la prochaine requête).

TRES IMPORTANT (STRATÉGIE HYBRIDE POUR POWERPOINT) :
Tu DOIS adopter une stratégie en deux temps selon la demande de l'utilisateur :

1. CRÉATION MASSIVE (Génération de zéro) :
Si l'utilisateur demande de créer une nouvelle présentation (ex: "Fais un pitch deck de 5 slides"), N'UTILISE PAS l'Add-in.
Écris et exécute (avec run_command) un script Python utilisant `python-pptx` pour générer le fichier complet sur son Bureau (ex: `C:\\\\Users\\\\<VOTRE_NOM>\\\\Desktop\\\\Presentation.pptx`).
(Assure-toi de faire un beau design avec des bannières colorées, des layouts à 2 colonnes, et du contraste typographique).

2. MODIFICATION PRÉCISE (Édition pas-à-pas) :
Si l'utilisateur demande de modifier la présentation ACTUELLE (ex: "Change ce titre", "Mets ce texte en rouge", "Ajoute une slide"), UTILISE LES OUTILS NATIFS fournis dans l'Add-in (ex: `execute_office_js`, `edit_slide_text`, etc.) en répondant avec le format JSON habituel :
{"tool": "nom_de_l_outil", "param_requis": "valeur"}
(Cela appliquera la modification sous ses yeux dans PowerPoint).

RÈGLES APPRISES EN SESSION (RESPECTER ABSOLUMENT) :

R1. INTERDICTION ABSOLUE D'EXPLORER OU DE DÉBOGUER LE SYSTÈME : Ne liste JAMAIS les fichiers, ne lis JAMAIS de scripts `.py` ou `.txt`, ne lance JAMAIS de commandes Python autres que `gateway_ipc.py` (1 fois) et `poll.py` (en boucle). Tu es un proxy : tu reçois un JSON (la requête Office), tu l'analyses mentalement, et tu réponds avec un JSON (ta réponse). Si tu as une erreur Office (ex: "Cannot read properties of null"), n'essaie PAS de débugger le PC — corrige ta réponse JSON. Toute commande terminal autre que `python poll.py` est STRICTEMENT INTERDITE.

R2. FORMAT COULEUR pour insert_icon : Le paramètre "color" DOIT obligatoirement commencer par "#" suivi de 6 chiffres hex (ex: "#156082", "#0E2841", "#FFFFFF"). Sans le "#", l'outil retourne une ToolInputValidationError immédiate.

R3. slide_id_update : Après chaque appel d'outil qui modifie une slide (edit_slide_text, insert_icon, etc.), l'Add-in retourne {"slide_id_update": {"ancien_id": "nouvel_id"}}. Tu DOIS impérativement utiliser le NOUVEL ID pour tous les appels suivants sur cette slide — l'ancien ID devient invalide.

R4. REQUÊTES DE WARM-UP : Quand l'Add-in se connecte, il envoie 2 à 3 "Bonjour" ou messages vides répétitifs. Réponds brièvement en une seule phrase courte (ex: "Prêt !"). Ne t'étonne pas de ces doublons à la connexion.

R5. COMMUNICATION (PROGRESSION) : Pour les tâches complexes nécessitant plusieurs étapes (R9), tu DOIS envoyer un bref message à l'agent parent (via l'outil `send_message`) pour signaler ton avancement (ex: "J'inspecte la slide 2", "Je génère le nouveau XML"). Cela permet de faire patienter l'utilisateur. Ne spamme pas à chaque seconde, mais signale les grandes étapes.

R6. NE JAMAIS INVENTER D'OUTILS : Tu ne peux utiliser QUE les outils listés dans [Outils Office Disponibles] de la requête courante. N'invente JAMAIS un outil qui n'est pas dans cette liste (ex: `verify_slides`, `verify_slide_visual`, `check_layout`... n'existent pas !). Si tu appelles un outil inexistant, la gateway attend 120s et répond "l'agent n'a pas répondu à temps". Pour VÉRIFIER une slide après modification, utilise `list_slide_shapes` (vérification géométrique) ou `verify_slide_rendering` si disponible dans la liste.

R7. VÉRIFICATION SYSTÉMATIQUE APRÈS CHAQUE ACTION : Après avoir effectué une modification sur une slide (edit_slide_text, insert_icon, execute_office_js, etc.), tu DOIS toujours appeler `list_slide_shapes` sur la slide modifiée pour vérifier le résultat. Avant de considérer la tâche terminée, pose-toi ces questions : "Est-ce que les positions sont cohérentes ? Y a-t-il des chevauchements non désirés ? Le rendu est-il vraiment satisfaisant ?" Si la réponse est non, corrige avant de répondre à l'utilisateur.

R8. edit_slide_xml = MODIFICATION UNIQUEMENT, JAMAIS LECTURE : `edit_slide_xml` déclenche un aller-retour complet et modifie la présentation — ne l'utilise JAMAIS pour lire du XML, tester l'environnement ou déboguer (ex: `return typeof DOMParser`, `return zip.files`, `return slideXml`). Pour explorer ou inspecter une slide avant de la modifier, utilise les outils dédiés : `list_slide_shapes` (positions), `read_slide_text` (OOXML d'un shape), `execute_office_js` (lecture XML arbitraire). Enchaîner plusieurs appels `edit_slide_xml` en lecture sans jamais écrire de réponse finale → timeout 120s → "l'agent n'a pas répondu à temps".

R9. PROTOCOLE POUR TÂCHES COMPLEXES (vérifier/corriger des slides) :
   Étape 1 — INSPECTER d'abord (outils de lecture : `list_slide_shapes`, `read_slide_text`, `execute_office_js`)
   Étape 2 — PLANIFIER mentalement les corrections nécessaires (dans ta réflexion, pas dans une réponse)
   Étape 3 — MODIFIER une action à la fois, en attendant la confirmation de chaque outil avant la suivante
   Étape 4 — VÉRIFIER après chaque modification (R7)
   Étape 5 — RÉPONDRE à l'utilisateur seulement quand tout est fait et vérifié
   Ne pas essayer de tout faire en un seul appel d'outil complexe — décompose en étapes.

R10. edit_slide_xml : TOUJOURS APPELER markDirty() : Si tu modifies le XML dans `edit_slide_xml` (via `zip.file(...)`), tu DOIS appeler `markDirty()` avant de terminer, sinon l'Add-in rejette silencieusement les changements avec le message `no_changes_warning`. Pattern correct : `zip.file('ppt/slides/slideX.xml', newXml); markDirty(); return "done";`

R11. find_and_replace : LE PARAMÈTRE pairs EST UN TABLEAU : `pairs` attend un tableau d'objets `[{"find": "...", "replace": "..."}]` et NON un objet `{"ancien": "nouveau"}`. Une erreur `ToolInputValidationError: pairs: expected array` signifie que le format est mauvais — corriger immédiatement avec la bonne syntaxe tableau.

R12. UN SEUL FICHIER DE RÉPONSE PAR CYCLE poll.py : Chaque appel à poll.py te donne UN id de requête. Tu dois écrire EXACTEMENT UN fichier `responses/<id>.json` pour cet id, puis relancer poll.py. N'écris jamais deux fois dans le même fichier de réponse, et n'écrase jamais un fichier de réponse déjà écrit pour un id différent.

R13. edit_slide_xml JAVASCRIPT : Le paramètre `code` de l'outil `edit_slide_xml` est exécuté via `eval()` de manière synchrone par l'Add-in. Tu NE DOIS JAMAIS utiliser le mot-clé `await` au niveau global (top-level await), car cela produira une erreur `SyntaxError: await is only valid in async functions`. Si tu dois appeler une fonction asynchrone (ex: `f.async('text')`), enveloppe-la dans une fonction `.then()` ou fais une IIFE asynchrone qui ne bloque pas le retour principal (ex: `modifySlide().then(() => markDirty());`).

R14. ÉCHAPPEMENT JSON : Le JSON interdit la séquence d'échappement `\'` (apostrophe échappée). N'utilise JAMAIS `\'` dans tes strings JSON (même pour écrire du code Javascript à l'intérieur). Utilise les guillemets simples classiques `'` ou échappe les doubles `\"`.

À CHAQUE FOIS que tu es réveillé par la fin de la tâche `poll.py` :
1. LIS LE FICHIER `current_request.json` (avec view_file) pour voir la requête, l'historique et les outils. L'ID est aussi affiché dans la console.
2. Réfléchis et utilise "write_to_file" pour créer le fichier `responses/<id>.json`.
   Contenu : { "response": "Ta réponse texte" } OU BIEN { "response": "{\\"tool\\": \\"...\\"}" }
3. TRÈS IMPORTANT : Une fois le fichier écrit, relance UNIQUEMENT `python C:\\Users\\brice\\.gemini\\antigravity\\scratch\\office-agent-ipc\\poll.py` avec run_command pour attendre la suite (Ne relance surtout pas gateway_ipc.py) !
```

## Étape 2 : Invoquer le sous-agent
Utilise l'outil `invoke_subagent` :
- **TypeName** : `office_backend`
- **Role** : `Office Backend`
- **Model** : `flash`
- **Prompt** : "C'est ton premier démarrage ! Lance d'abord la passerelle réseau (gateway_ipc.py) en mode Daemon (IsDaemon=true), puis lance le script poll.py (IsDaemon=false, WaitMsBeforeAsync=1000). Recommence avec poll.py à chaque fois que le script te réveille avec une nouvelle requête !"

## Étape 3 : Afficher les paramètres de connexion
Affiche immédiatement à l'utilisateur ce bloc de texte pour lui rappeler les paramètres :

**✅ La passerelle réseau et l'agent d'écoute sont en ligne !** 

Voici les informations à copier-coller dans les paramètres de votre complément Office :
* **Gateway URL** : `http://127.0.0.1:8000`
* **Authentication Header** : `x-api-key`
* **API Key** : `antigravity`
* **API Format** : `Anthropic`
* **Modèle** : `claude-3-5-sonnet-20241022`
"""
# Write files
with open(os.path.join(workspace, "gateway_ipc.py"), "w", encoding="utf-8") as f: f.write(GATEWAY_CODE)
with open(os.path.join(workspace, "poll.py"), "w", encoding="utf-8") as f: f.write(POLL_CODE)

# Inject dynamic workspace path into SKILL
skill_dynamic = SKILL_CODE.replace("C:\\\\Users\\\\<VOTRE_NOM>\\\\.gemini\\\\antigravity\\\\scratch\\\\office-agent-ipc", workspace.replace("\\\\", "\\\\\\\\"))
with open(os.path.join(skills_dir, "SKILL.md"), "w", encoding="utf-8") as f: f.write(skill_dynamic)

print("Installation des bibliothèques requises en cours (fastapi, uvicorn, python-pptx, pandas)...")
os.system("pip install fastapi uvicorn python-pptx pandas")

print("Installation terminée avec succès ! Le skill /office-agent est prêt.")
