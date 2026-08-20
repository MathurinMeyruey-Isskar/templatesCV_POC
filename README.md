# 🚀 Mon Projet Streamlit avec UV

Ce dépôt contient une application **Streamlit** écrite en Python, gérée avec [**UV**](https://github.com/astral-sh/uv) pour la gestion des dépendances et des environnements.

---

## 📌 Prérequis

- **Python 3.11 ou supérieur** (recommandé pour UV)
- **UV** (pour gérer les dépendances et exécuter l'application)

---

## 🛠 Installation

### 1. Cloner le dépôt

```bash
git clone https://github.com/ton-utilisateur/mon-projet.git
cd mon-projet
```

### 2. Installer UV

Si ce n'est pas déjà fait, installe **UV** en suivant les instructions officielles :

#### **Sur Linux/macOS**

```bash
curl -LsSf https://astral-sh.github.io/uv/install.sh | sh
```

#### **Sur Windows (PowerShell)**

```powershell
irm https://astral-sh.github.io/uv/install.ps1 | iex
```

> ⚠️ Assure-toi que le chemin d'installation de UV est dans ton `PATH` (ex: `~/.local/bin` sur Linux/macOS).

### 3. Créer un environnement virtuel et installer les dépendances

UV gère automatiquement les environnements virtuels. Pour installer les dépendances du projet :

```bash
uv sync
```

> Cette commande crée un environnement virtuel (`.venv`) et installe toutes les dépendances listées dans le fichier `pyproject.toml` ou `requirements.txt`.

---

## 🏃‍♂️ Lancer l'application

Pour exécuter l'application Streamlit :

```bash
uv run streamlit run app.py
```

> L'application sera accessible dans ton navigateur à l'adresse :  
> [**http://localhost:8501**](http://localhost:8501)

---

## 📂 Structure du projet

```
mon-projet/
├── app.py              # Point d'entrée de l'application Streamlit
├── pyproject.toml      # Configuration des dépendances (si utilisée)
├── requirements.txt    # Dépendances Python (alternative)
└── README.md           # Ce fichier
```

---

## 🤝 Contribution

1. Fork ce dépôt.
2. Crée une branche pour ta fonctionnalité (`git checkout -b ma-fonctionnalité`).
3. Commit tes modifications (`git commit -m "Ajout de ma fonctionnalité"`).
4. Pousse vers la branche (`git push origin ma-fonctionnalité`).
5. Ouvre une **Pull Request**.

---
