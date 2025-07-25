# 🔗 Path Finder for Outlook (macOS) 💻

Ce projet vous permet de convertir des chemins de fichiers macOS ou windows dans les e-mails d'Outlook Web en liens cliquables qui déclenchent Finder pour ouvrir le dossier correspondant. Il utilise un serveur Python et un script Tampermonkey pour une intégration fluide dans l'interface web d'Outlook.

## 📋 À faire

- ✅ Empêcher l'ouverture d'un nouvel onglet si l'onglet est déjà ouvert dans un finder, et donc l'afficher à la place.

--- 

## 🌟 Fonctionnalités

- 🔄 Convertit automatiquement les chemins de fichiers macOS (par exemple, `/Volumes/...`) en liens cliquables dans Outlook Web.
- 📂 En cliquant sur le lien, le dossier spécifié s'ouvre dans Finder.
- 🚪 Ferme automatiquement les onglets ouverts par le processus (en utilisant un onglet temporaire pour gérer les requêtes du serveur local).

# 🎬 Demo

![demo](https://github.com/user-attachments/assets/69a80022-3cfa-4dda-9935-e34cd2f6c318)
![demo](https://github.com/TWhiteShadow/UniversalPathConverter/blob/main/demo.gif?raw=true)


---
## 🔧 Prérequis

Avant de commencer, assurez-vous d'avoir les éléments suivants installés :

- 🐍 [Python 3.x](https://www.python.org/downloads/)
- 🛠️ [Tampermonkey](https://www.tampermonkey.net/) (une extension de navigateur pour exécuter des scripts utilisateur)

---

## ⚙️ Instructions d'installation

### 1. 🛠️ Configuration du Serveur Python

Le serveur Python est chargé d’ouvrir Finder sur macOS lorsque le lien est cliqué. Il écoute sur `http://localhost:6969` et utilise AppleScript pour interagir avec Finder.

#### ► Étape 1 : Créer le Script du Serveur Python

Créez un fichier nommé `path_opener_server.py` et ajoutez le code suivant :

```python
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, unquote
import subprocess
from typing import Optional
import logging
import os

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

class PathOpenerHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        """Handle GET requests by opening the specified path in Finder"""
        try:
            # Parse the URL and extract the path parameter
            parsed_url = urlparse(self.path)
            if parsed_url.path != '/open':
                self.send_error(404, "Path not found")
                return

            # Get the path query parameter and decode it
            query_dict = dict(pair.split('=') for pair in parsed_url.query.split('&'))
            encoded_path = query_dict.get('path')
            
            if not encoded_path:
                self.send_error(400, "No path parameter provided")
                return
            
            # Decode the URL-encoded path
            mac_path = unquote(encoded_path)
            
            # Verify the path exists
            if not os.path.exists(mac_path):
                error_msg = f"Path does not exist: {mac_path}"
                logger.error(error_msg)
                self.send_error(404, error_msg)
                return
            
            # Open the path in Finder
            self._open_in_finder(mac_path)
            
            # Send success response
            self.send_response(200)
            self.send_header('Content-type', 'text/plain')
            self.send_header('Access-Control-Allow-Origin', '*')  # Allow CORS
            self.end_headers()
            self.wfile.write(f"Opening path: {mac_path}".encode())
            
            logger.info(f"Successfully opened path: {mac_path}")
            
        except Exception as e:
            logger.error(f"Error processing request: {str(e)}")
            self.send_error(500, f"Internal server error: {str(e)}")
    
    def _open_in_finder(self, path: str) -> Optional[subprocess.CompletedProcess]:
        """
        Open the specified path in Finder using AppleScript
        """
        try:
            # First check if Finder is running and has any windows open
            check_windows_script = '''
                tell application "Finder"
                    if (count of windows) is greater than 0 then
                        return "true"
                    else
                        return "false"
                    end if
                end tell
            '''
            
            check_result = subprocess.run(['osascript', '-e', check_windows_script], 
                                        capture_output=True, 
                                        text=True)
            
            has_windows = check_result.stdout.strip().lower() == "true"
            
            if has_windows:
                # If Finder has windows, create a new tab and set it to the path
                new_tab_script = f'''
                    tell application "Finder"
                        activate
                        set myFile to POSIX file "{path}" as alias
                        set parentFolder to container of myFile
                    end tell
                    tell application "System Events"
                        keystroke "t" using command down
                        delay 0.1  -- Brief pause to ensure the new tab is created
                    end tell
                    tell application "Finder"
                        tell front window
                            set target to parentFolder
                            select myFile
                        end tell
                    end tell
                '''
            else:
                # If no windows exist, reveal the item which will open a new window
                new_tab_script = f'''
                    tell application "Finder"
                        reveal (POSIX file "{path}" as alias)
                        activate
                    end tell
                '''
            
            result = subprocess.run(['osascript', '-e', new_tab_script], 
                                    capture_output=True, 
                                    text=True)
            
            if result.returncode != 0:
                logger.error(f"AppleScript error: {result.stderr}")
                return None
            
            return result
            
        except subprocess.SubprocessError as e:
            logger.error(f"Failed to execute commands: {str(e)}")
            return None


def run_server(port: int = 6969):
    """Start the HTTP server"""
    server_address = ('', port)
    httpd = HTTPServer(server_address, PathOpenerHandler)
    logger.info(f"Server starting on port {port}...")
    
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        logger.info("Server stopping...")
        httpd.server_close()

if __name__ == "__main__":
    run_server()
```

#### ► Étape 2 : Exécuter le Serveur Python

Lancez le serveur en utilisant la commande suivante :

```bash
python3.10 path_opener_server.py
```
ou
```bash
python path_opener_server.py
```

💡 Le serveur commencera à écouter sur `http://localhost:6969` pour les requêtes.

---

### 2. 🚀 Configuration du Lancement Automatique sur macOS

Pour que le serveur Python démarre automatiquement lors de la connexion à votre Mac, vous pouvez utiliser la méthode suivante :

1. 📝 Créez un fichier de configuration LaunchAgent :

```bash
mkdir -p ~/Library/LaunchAgents
code ~/Library/LaunchAgents/com.raph.universalpathconverter.server.plist
```

2. 📄 Collez le contenu suivant dans le fichier :

```xml
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>com.raph.universalpathconverter.server</string>

    <key>ProgramArguments</key>
    <array>
        <string>/usr/bin/python3</string>
        <string>/Users/raphaeltoursel/Projects/finder_opener_works.py</string>
    </array>

    <key>RunAtLoad</key>
    <true/>
    <key>KeepAlive</key>
    <true/>
</dict>
</plist>
```

3. ▶️ Chargez le service :

```bash
launchctl load ~/Library/LaunchAgents/raph.universalpathconverter.server.plist
```

---

### 3. 🖱️ Configuration du Script Tampermonkey

Ce script Tampermonkey modifiera l’application Web Outlook pour remplacer les chemins de fichiers macOS par des liens cliquables. Lorsqu’ils sont cliqués, ils envoient une requête au serveur Python qui ouvre le dossier dans Finder.

#### ► Étape 1 : Installer Tampermonkey

Installez Tampermonkey pour votre navigateur depuis le [site officiel](https://www.tampermonkey.net/).

#### ► Étape 2 : Créer un Nouveau Script Tampermonkey

Après avoir installé Tampermonkey, créez un nouveau script utilisateur et collez le code suivant dans l’éditeur :

```javascript
// ==UserScript==
// @name         Mac Path Finder Linker for Outlook (macOS) with Local Server and Auto-Close
// @namespace    http://tampermonkey.net/
// @version      1.6
// @description  Convertit les chemins de fichiers macOS et Windows dans les e-mails Outlook en liens cliquables qui déclenchent un serveur local pour ouvrir Finder
// @match        https://outlook.office.com/mail/*
// @grant        none
// @match        http://localhost:6969/*
// ==/UserScript==

(function() {
    'use strict';

    // Fonction pour fermer l'onglet si l'URL est localhost:6969
    function closeIfLocalhost() {
        if (window.location.href.startsWith('http://localhost:6969')) {
            window.close();
        }
    }

    // Vérifie si l'URL actuelle est localhost:6969 et ferme l'onglet
    closeIfLocalhost();

    // Fonction pour convertir un chemin Windows en chemin macOS
    function convertWindowsToMacPath(windowsPath) {
        // Remplace H:\ par /Volumes/HYUNDAI/ et convertit les backslashes en forward slashes
        return windowsPath.replace(/^H:\\/, '/Volumes/HYUNDAI/').replace(/\\/g, '/');
    }

    // Fonction pour convertir les chemins de fichiers en liens cliquables dans les e-mails Outlook
    function convertPathsToLinks() {
        console.log('Script en cours d\'exécution...'); // Journal de débogage

        // Récupère tous les éléments contenant du texte
        const elements = document.querySelectorAll('[role="main"] *');

        elements.forEach(element => {
            // Ignore si l'élément n'a pas de contenu texte ou est déjà un lien
            if (!element.childNodes || element.tagName === 'A') return;

            // Traite uniquement les nœuds de texte
            element.childNodes.forEach(node => {
                if (node.nodeType === Node.TEXT_NODE) {
                    const text = node.textContent;
                    if (text.includes('/Volumes/') || text.includes('H:\\')) {
                        console.log('Chemin trouvé :', text); // Journal de débogage

                        // Crée un nouvel élément span avec le chemin lié
                        const span = document.createElement('span');
                        span.innerHTML = text.replace(
                            /(\/Volumes\/[^\s<>"]*|H:\\[^\s<>"]*)/g,
                            match => {
                                // Convertit le chemin Windows en chemin macOS si nécessaire
                                const macPath = match.startsWith('H:\\') ? convertWindowsToMacPath(match) : match;

                                // Encode le chemin de fichier pour le serveur local
                                const filePath = encodeURIComponent(macPath);

                                // Crée le lien pour déclencher le serveur local
                                const localServerUrl = `http://localhost:6969/open?path=${filePath}`;

                                console.log('Création du lien pour le serveur local :', localServerUrl); // Journal de débogage

                                return `<a href="${localServerUrl}" target="_blank"
                                          style="color: #0078d4; text-decoration: underline; cursor: pointer;"
                                          onclick="event.preventDefault();
                                                   const newTab = window.open('${localServerUrl}', '_blank');
                                                   setTimeout(() => { newTab.close(); }, 1000);"
                                      >${match}</a>`;
                            }
                        );

                        // Remplace le nœud de texte par notre nouveau span
                        node.parentNode.replaceChild(span, node);
                    }
                }
            });
        });
    }

    // Exécute au chargement de la page avec un délai pour assurer le chargement d'Outlook
    setTimeout(() => {
        console.log('Exécution initiale...'); // Journal de débogage
        convertPathsToLinks();
    }, 2000);

    // Exécute lorsque de nouveaux e-mails sont chargés
    const observer = new MutationObserver((mutations) => {
        console.log('Changement de contenu, exécution du convertisseur...'); // Journal de débogage
        convertPathsToLinks();
    });

    // Commence l'observation
    setTimeout(() => {
        const mainContent = document.querySelector('[role="main"]');
        if (mainContent) {
            observer.observe(mainContent, {
                childList: true,
                subtree: true
            });
            console.log('Observateur démarré'); // Journal de débogage
        }
    }, 2000);
})();
```

#### ► Étape 3 : Sauvegardez et Activez le Script

- 💾 Sauvegardez le script dans Tampermonkey.
- ✅ Assurez-vous que le script est activé.

---

## 🎬 Fonctionnement

1. **⚙️ Serveur Python** : Le serveur Python écoute sur `http://localhost:6969` et utilise AppleScript pour ouvrir Finder lorsqu'un lien est cliqué.
2. **🖱️ Script Tampermonkey** : Le script Tampermonkey modifie l’application Web Outlook pour détecter les chemins de fichiers macOS, les remplacer par des liens cliquables, et envoyer une requête au serveur Python pour ouvrir Finder.
3. **❌ Fermeture Automatique des Onglets** : Le script ferme également tout onglet temporaire ouvert par le lien vers `localhost:6969`, évitant ainsi de montrer des fenêtres inutiles à l’utilisateur.

---

## 🔍 Utilisation

1. 🌐 Ouvrez Outlook Web dans votre navigateur.
2. 📁 Tout e-mail contenant des chemins de fichiers macOS (commençant par `/Volumes/`) aura désormais des liens cliquables.
3. 🖱️ Cliquez sur le lien pour ouvrir le dossier dans Finder.

---

### 🛠️ Dépannage

- **❗ Serveur Python Non Démarré** : Assurez-vous que le serveur Python est en cours d'exécution et écoute sur le port 6969.
- **❗ Onglets Non Fermés** : Si les onglets ne se ferment pas automatiquement après avoir cliqué sur le lien, assurez-vous que le script Tampermonkey est activé et que les paramètres de votre navigateur permettent la fermeture d'onglets par programme.


---

