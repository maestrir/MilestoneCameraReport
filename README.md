
# 📹 Milestone Camera Report Script (PowerShell)

Questo script PowerShell crea un **report completo** delle telecamere gestite da Milestone XProtect, con informazioni dettagliate sullo stato operativo, dati di registrazione e snapshot visivi integrati in un file Excel.

## 🚀 Funzionalità principali

- **Dati dettagliati** di ogni telecamera, inclusi:
  - Nome telecamera
  - Stato (abilitata/disabilitata)
  - Spazio utilizzato per le registrazioni
  - Data inizio e fine registrazione
  - Giorni effettivi di conservazione video
  - Descrizione hardware e indirizzo IP
- **Immagine snapshot** dell’ultimo fotogramma registrato, integrata direttamente nel report Excel per una visualizzazione immediata.

## 🛠️ Prerequisiti

Prima di utilizzare lo script, installare questi moduli PowerShell:

```powershell
Install-Module MilestonePSTools, ImportExcel -Scope CurrentUser
```

## 🔔 Informazioni importanti

**⚠️ Nota sulle telecamere disabilitate:**  
Solo le telecamere **attive e abilitate** mostrano dati completi come spazio utilizzato, periodo di registrazione e snapshot.

Se si desidera un report completo **anche per le telecamere disabilitate**, è necessario temporaneamente attivarle. Dopo la generazione del report, ricorda di **disabilitarle nuovamente** per evitare conseguenze indesiderate sul consumo delle licenze Milestone.

## ▶️ Come utilizzare lo script

1. Aprire PowerShell con privilegi adeguati.
2. Collegarsi al server Milestone tramite lo script (apparirà una finestra di login).
3. Eseguire lo script.
4. Attendere il completamento: il report Excel finale verrà generato nella cartella dello script.

**Esempio d'uso:**

```powershell
.\MilestoneCameraReport.ps1
```

## 📌 Output generato

Il report Excel finale con snapshot integrati verrà salvato automaticamente nella cartella dove risiede lo script, con nome composto da:

```
RecorderName_ReportTelecamere_<data-ora>.xlsx
```

📂 **Una sottocartella** contenente le immagini snapshot verrà generata automaticamente nella stessa posizione.

## 🚩 Licenza

Questo script è reso disponibile gratuitamente ed è utilizzabile liberamente. Modificalo in base alle tue esigenze.

---

✅ **Ideale per audit, documentazione e monitoraggio periodico di sistemi di videosorveglianza Milestone XProtect.**

