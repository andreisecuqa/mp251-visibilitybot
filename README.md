# Bot Telegram — Examen Managementul Vizibilității

Generează automat proiectul de examen (.xlsx completat) pe baza organizației alese de fiecare student.

---

## 📦 Instalare

```bash
pip install -r requirements.txt
```

---

## ⚙️ Configurare

Ai nevoie de 2 chei:

### 1. Token Telegram Bot
1. Deschide Telegram → caută `@BotFather`
2. Scrie `/newbot` → alege un nume și un username
3. Copiază tokenul (format: `123456:ABC-DEF...`)

### 2. Cheie Gemini API
1. Du-te la https://aistudio.google.com/app/apikey
2. Click "Create API Key"
3. Copiază cheia

---

## 🚀 Pornire

### Varianta 1 — Variabile de mediu (recomandat)
```bash
export TELEGRAM_TOKEN="tokenul_tau_aici"
export GEMINI_API_KEY="cheia_gemini_aici"
python bot.py
```

### Varianta 2 — Editează direct în cod
Deschide `bot.py` și înlocuiește:
```python
TELEGRAM_TOKEN = "YOUR_BOT_TOKEN"   # linia ~17
GEMINI_API_KEY = "YOUR_GEMINI_KEY"  # linia ~18
```

---

## 📁 Structura proiectului

```
exam_bot/
├── bot.py              # Codul principal al botului
├── template.xlsx       # Fișierul Excel original (template)
├── requirements.txt    # Dependențe Python
├── README.md           # Acest fișier
└── outputs/            # Fișierele generate (se creează automat)
```

---

## 💬 Cum folosesc studenții botul

1. Deschid Telegram → caută botul după username
2. Scriu `/start`
3. Răspund la 8 întrebări (organizație, domeniu, platforme etc.)
4. Confirmă datele
5. Primesc fișierul `.xlsx` complet în ~60-90 secunde

---

## ⚠️ Ce trebuie făcut manual după generare

Botul completează **~80% din proiect**. Rămân manuale:
- Inserarea screenshot-urilor reale din paginile sociale
- Linkurile reale ale postărilor performante
- Crearea celor 4 materiale în Canva + inserarea lor în Excel
- Verificarea și personalizarea conținutului generat

---

## 🔧 Găzduire (optional — pentru acces 24/7)

Poți rula botul pe:
- **PC personal** — simplu, dar se oprește când închizi laptopul
- **Railway.app** — gratuit, deploy în 5 minute (recomandat)
- **Render.com** — alternativă gratuită
- **VPS** — pentru grupe mari

### Deploy pe Railway:
1. Creează cont pe railway.app
2. New Project → Deploy from GitHub
3. Adaugă variabilele de mediu: `TELEGRAM_TOKEN` și `GEMINI_API_KEY`
4. Start command: `python bot.py`

---

## 📊 Costuri estimate

- **Telegram Bot API**: GRATUIT
- **Gemini Flash API**: GRATUIT (limita generoasă: 15 req/min, 1500 req/zi)
- Pentru o grupă de 30 studenți: 0 lei

---

## 🛟 Troubleshooting

**Eroare: `json.JSONDecodeError`**
→ Gemini uneori returnează text extra. Botul încearcă să curețe automat. Dacă persistă, rulează din nou /start.

**Eroare: `TELEGRAM_TOKEN not found`**
→ Asigură-te că ai setat variabila de mediu corect.

**Botul nu răspunde**
→ Verifică că `python bot.py` rulează în terminal fără erori.
