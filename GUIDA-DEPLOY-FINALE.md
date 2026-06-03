# Il Sipario Musicale — Guida Deploy Produzione

## Panoramica
Questa guida ti porta dall'avere i file a un sito funzionante con login, database condiviso, e tutte le funzionalità.

**Tempo stimato: 30-40 minuti**

---

## FASE 1: Supabase (Database + Login)

### 1.1 Crea il progetto Supabase (3 min)
1. Vai su **https://supabase.com** e clicca "Start your project"
2. Registrati con GitHub (usa lo stesso account di prima)
3. Clicca "New project"
4. Compila:
   - **Name**: `sipario-preventivi`
   - **Database Password**: scegli una password sicura (salvala da qualche parte)
   - **Region**: West EU (London) — il più vicino all'Italia
5. Clicca "Create new project" — aspetta ~2 minuti

### 1.2 Crea le tabelle (2 min)
1. Nel pannello Supabase, vai su **SQL Editor** (icona nel menu laterale)
2. Clicca "New query"
3. Apri il file `supabase/schema.sql` dal progetto
4. Copia TUTTO il contenuto e incollalo nel query editor
5. Clicca **"Run"** (bottone verde)
6. Deve apparire "Success"

### 1.3 Prendi le chiavi (1 min)
1. Vai su **Settings → API** (nel menu laterale)
2. Copia questi due valori:
   - **Project URL**: qualcosa tipo `https://xxxxx.supabase.co`
   - **anon public key**: una stringa lunga tipo `eyJhbGc...`
3. Tienili pronti per dopo

### 1.4 Crea il primo utente admin (2 min)
1. Vai su **Authentication → Users** (nel menu laterale)
2. Clicca "Add user" → "Create new user"
3. Inserisci:
   - Email: l'email dell'admin (es. fabrizio@sipariomusicale.com)
   - Password: una password
   - Spunta "Auto Confirm User"
4. Clicca "Create user"
5. Ora vai su **Table Editor → profiles**
6. Trova l'utente appena creato
7. Clicca sulla riga, cambia il campo `role` da `operator` a `admin`
8. Clicca "Save"

L'admin potrà poi creare altri utenti dall'app.

---

## FASE 2: Deploy su Vercel

### 2.1 Aggiorna i file su GitHub (3 min)
1. Vai nel tuo repository GitHub (`sipario-preventivi`)
2. Cancella tutti i file vecchi se ce ne sono
3. Scarica e decomprimi il nuovo `sipario-final.zip`
4. Carica tutti i file su GitHub (trascina nella pagina di upload)
5. Commit

### 2.2 Configura le variabili d'ambiente su Vercel (2 min)
1. Vai su **vercel.com** → il tuo progetto
2. Clicca **Settings → Environment Variables**
3. Aggiungi queste due variabili:

| Nome | Valore |
|------|--------|
| `VITE_SUPABASE_URL` | L'URL del progetto Supabase (copiato prima) |
| `VITE_SUPABASE_ANON_KEY` | La anon key (copiata prima) |

4. Clicca "Save" per ciascuna

### 2.3 Ri-deploy (1 min)
1. Vai su **Deployments** nel tuo progetto Vercel
2. Clicca sui tre puntini dell'ultimo deploy → "Redeploy"
3. Aspetta ~60 secondi

### 2.4 Fatto!
Vai all'URL del tuo sito e accedi con l'email/password dell'admin creato al passo 1.4.

---

## FASE 3: Creare gli altri utenti

Dall'app (loggato come admin):
1. Clicca "Gestione Utenti" nella homepage
2. Inserisci nome, email, password del nuovo operatore
3. Clicca "Crea Utente"
4. L'utente riceverà un'email di conferma (deve cliccare il link)
5. Dopo la conferma può accedere all'app

**Alternativa più veloce**: crea gli utenti direttamente da Supabase Dashboard (Authentication → Users → Add User con "Auto Confirm") come hai fatto per l'admin.

---

## Struttura dei File

```
sipario-final/
├── index.html          — Pagina HTML principale
├── package.json        — Dipendenze (React, Supabase, SheetJS)
├── vite.config.js      — Config build tool
├── .env.example        — Template variabili d'ambiente
├── src/
│   ├── main.jsx        — Entry point React
│   ├── App.jsx         — App completa (UI + logica)
│   ├── supabase.js     — Client Supabase
│   └── excel.js        — Generazione Excel e PDF
└── supabase/
    └── schema.sql      — Schema database (da eseguire una volta)
```

---

## Aggiornamenti Futuri

Per aggiornare l'app:
1. Modifica i file in `src/`
2. Carica su GitHub
3. Vercel ri-deploya automaticamente in 30 secondi
4. I dati nel database NON vengono toccati

Per aggiungere campi al database:
1. Scrivi un nuovo script SQL
2. Eseguilo nel SQL Editor di Supabase
3. I dati vecchi restano intatti

---

## Costi

| Servizio | Costo |
|----------|-------|
| Supabase Free | €0/mese (fino a 50.000 righe) |
| Vercel Free | €0/mese |
| Dominio (opzionale) | ~€10/anno |
| **Totale** | **€0/mese** |

---

## Supporto

Se qualcosa non funziona:
- Controlla che le variabili d'ambiente su Vercel siano corrette
- Controlla che lo schema SQL sia stato eseguito senza errori
- Verifica in Supabase Dashboard → Authentication che gli utenti esistano
- Controlla la console del browser (F12 → Console) per errori
