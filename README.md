# 🚀 Antigravity ↔ Office Agent (Installation & Guide)

Bienvenue dans l'outil d'intégration locale d'Antigravity avec la suite Microsoft Office (Word, Excel, PowerPoint) ! 
Ce dossier contient tout ce qu'il faut pour transformer n'importe quel ordinateur équipé d'Antigravity et d'Office en un assistant IA surpuissant, **sans aucune clé API payante**, et tournant **100% localement**.

---

## 🛠️ 1. Installation (Une seule fois par PC)

1. Assurez-vous d'avoir **Python** installé sur votre machine.
2. Double-cliquez sur le fichier **`Installer.bat`** présent dans ce dossier.
3. C'est fait ! Le script va automatiquement créer les composants et installer la compétence `/office-agent` dans Antigravity.

---

## 🔄 2. Utilisation Quotidienne (De A à Z)

À chaque fois que vous souhaitez utiliser l'agent dans Word, Excel ou PowerPoint, c'est **totalement automatique** :

### Étape A : Lancer l'Agent Fantôme (Et le Réseau)
1. Ouvrez votre application de messagerie **Antigravity**.
2. Dans une nouvelle discussion, tapez simplement la commande magique :
   `/office-agent`
3. Magie ! 🪄 L'agent va automatiquement **allumer le serveur réseau en arrière-plan** puis se mettre sur écoute. 
4. Quand il vous confirme qu'il écoute, vous pouvez réduire la fenêtre. *(Plus besoin de script .bat sur le bureau !)*

### Étape B : Connecter Word / Excel / PowerPoint
1. Ouvrez votre fichier (Word, Excel ou PPT).
2. Lancez le complément (Add-in) **Claude for Office**.
3. Dans la roue crantée des paramètres (Settings), configurez la connexion ainsi :
   - **Gateway URL** : `http://127.0.0.1:8000` (ou `http://127.0.0.1:8000/v1`)
   - **Authentication Header** : `x-api-key`
   - **API Key** : `antigravity` (ce mot de passe est factice, tapez ce que vous voulez).
   - **API Format** : `Anthropic`
4. Cliquez sur **Connect**. Si le voyant passe au vert, vous pouvez commencer à discuter !

---

## 💡 3. Ce que l'agent peut faire

Grâce à notre nouvelle architecture dynamique, **l'agent s'adapte automatiquement au logiciel que vous utilisez !**
Quand vous lui parlez, le complément lui envoie discrètement la liste des outils disponibles (lire un paragraphe dans Word, écrire dans une cellule Excel, ajouter une diapo PPT...). 

Vous pouvez discuter normalement avec lui :
- *"Analyse cette ligne"* ➡️ L'agent va lire la cellule Excel de lui-même en arrière-plan.
- *"Résume ce document"* ➡️ L'agent va utiliser l'outil de lecture Word.
- *"Remplace le premier paragraphe par ceci..."* ➡️ L'agent utilisera l'outil de modification de Word.

*Note : La toute première question peut prendre quelques secondes de plus (le temps de faire l'aller-retour invisible avec le logiciel pour utiliser l'outil), c'est normal !*

---

🎉 **Bon travail !**
*(Créé par votre assistant IA Antigravity)*
