# mdaemon-mail-counter

![Python 3.8+](https://img.shields.io/badge/Python-3.8%2B-3776AB?logo=python&logoColor=white)
![MDaemon logs](https://img.shields.io/badge/MDaemon-logs-C8102E)
![openpyxl](https://img.shields.io/badge/dependency-openpyxl-2E7D32)
![Excel report](https://img.shields.io/badge/output-Excel%20.xlsx-217346)

[English](#english) | [Italiano](#italiano)

## English

Python script that counts, for each account of a domain, the mail **sent** and **received** by reading MDaemon logs, and writes an Excel file with summary, details and checks.

### Requirements

- Python 3.8 or later
- [openpyxl](https://pypi.org/project/openpyxl/): `python -m pip install openpyxl`

### Usage

Put `mail_counter.py` in the folder with the MDaemon logs (loose files, or the log zips as MDaemon archives them) and run:

```
python mail_counter.py --domain example.com                                    # latest day found
python mail_counter.py --domain example.com --date 2026-01-15                  # one day
python mail_counter.py --domain example.com --from 2026-01-10 --to 2026-01-15  # date range
python mail_counter.py --domain example.com --date 2026-01-15 --all-log        # only the all.log file
```

Other options: `--folder` (log folder, default the script's folder), `--output` (Excel file), `--aliases`, `--mailboxes` and `--log-names` (configuration files), `--draft-mailboxes FILE` (writes a draft `mailboxes.txt` inferred from the logs). `python mail_counter.py -h` shows everything.

### Files read

For each day: `MDaemon-YYYY-MM-DD-Routing.log` (required), `-SMTP-out.log`, `-SMTP-in.log` and, optionally, the content filter log (`-Content-Filter.log` or `-ContentFilter.log`). These are the names written by an **English** MDaemon. MDaemon translates these names into the language it is installed in: the names for each language are listed in `log_names.txt` and the script tries them all, so the Italian names (`-Instradamento.log`, `-SMTP-(uscita).log`, `-SMTP-(entrata).log`, `-Filtro-contenuti.log`) are recognised automatically. Zips are read in memory, nothing is extracted to disk. If the same log is found more than once, the loose file wins, then the largest copy.

With `--all-log` the script reads only `MDaemon-YYYY-MM-DD-all.log`, the log where MDaemon writes everything, and ignores the other logs. That file can be hundreds of MB, because it also holds all the IMAP traffic: the script shows its size and asks for confirmation before reading it (`-y` or `--yes` skips the question; without a terminal and without `-y` it stops and does nothing). It is read in chunks, also from inside a zip, and gives the same counts as the separate logs.

### Configuration (optional)

- `aliases.txt`: `alias@domain = primary@domain`, one line per alias.
- `mailboxes.txt`: `primary@domain = C:\MDaemon\Users\domain\folder`, primary addresses only. Needed when the mailbox folder name does not match the address.
- `log_names.txt`: log file names for each MDaemon language, one section per language. **English** and **Italian** are filled in; German, French, Spanish, Portuguese, Russian, Japanese, Korean and Chinese are empty placeholders, because no reliable source was found for their names. If your logs are not found, uncomment the section of your language and fill it in with the names you see in your MDaemon `Logs` folder. The file is looked up next to the script or in the log folder (`--log-names FILE` to use another one); without it only the English names are used.

See `aliases.example.txt` and `mailboxes.example.txt`. Anything not configured is inferred from the logs and reported in the "Mailboxes and aliases" sheet, together with any conflict between configuration and logs.

### Counting method

- Each message that entered the server (INBOUND) is counted once; duplicates with the same Message-ID and sender are merged.
- **Sent**: messages whose sender belongs to the domain, split into internal only and with external recipients (real recipients taken from the SMTP out log). Auto-replies and server notifications are kept apart.
- **Received**: deliveries into mailboxes, once per message and mailbox, split into internal senders, external senders and server notifications.
- With a date range the logs are analysed together, so a message crossing midnight stays one message.
- **Size**: the size of each message is the one written in the routing log when the message enters the server. A message counts for its sender and for each mailbox that receives it. Messages without a size in the logs (some server notifications) are left empty and are not counted in totals and averages; their number is reported on screen and in the "Sources and checks" sheet.

### Excel output

By default the file only contains the **Summary by account** sheet, with daily averages and, for each account, the total size sent and received (MB) and the average size per message sent and received (KB). The **Message detail** and **Detail by account** sheets have a size column (KB), and **By day** has the MB sent and received per day. Add the other sheets with short flags, which can be combined:

| Flag | Sheet |
|------|-------|
| `-d` | By day (only with more than one day) |
| `-md` | Message detail |
| `-ad` | Detail by account |
| `-mb` | Mailboxes and aliases |
| `-s` | Sources and checks |
| `-a` | all sheets |

Example: `python mail_counter.py --domain example.com --from 2026-01-10 --to 2026-01-15 -d -md`

## Italiano

Script Python che conta, per ogni account di un dominio, le mail **inviate** e **ricevute** leggendo i log di MDaemon, e produce un file Excel con riepilogo, dettaglio e verifiche. Lo script e i fogli Excel sono in inglese.

### Requisiti

- Python 3.8 o successivo
- [openpyxl](https://pypi.org/project/openpyxl/): `python -m pip install openpyxl`

### Uso

Metti `mail_counter.py` nella cartella che contiene i log di MDaemon (file sciolti o gli zip dei log così come vengono archiviati) e lancia:

```
python mail_counter.py --domain example.com                                    # ultima giornata trovata
python mail_counter.py --domain example.com --date 2026-01-15                  # una giornata
python mail_counter.py --domain example.com --from 2026-01-10 --to 2026-01-15  # intervallo
python mail_counter.py --domain example.com --date 2026-01-15 --all-log        # solo il file all.log
```

Altre opzioni: `--folder` (cartella dei log, default quella dello script), `--output` (file Excel), `--aliases`, `--mailboxes` e `--log-names` (file di configurazione), `--draft-mailboxes FILE` (genera una bozza di `mailboxes.txt` ricavata dai log). `python mail_counter.py -h` mostra tutto.

### File letti

Per ogni giornata: `MDaemon-AAAA-MM-GG-Routing.log` (obbligatorio), `-SMTP-out.log`, `-SMTP-in.log` e, facoltativo, il log del filtro contenuti (`-Content-Filter.log` o `-ContentFilter.log`). Sono i nomi dei log di MDaemon in **inglese**. MDaemon traduce questi nomi nella lingua in cui è installato: i nomi per ogni lingua sono elencati in `log_names.txt` e lo script li prova tutti, quindi i nomi italiani (`-Instradamento.log`, `-SMTP-(uscita).log`, `-SMTP-(entrata).log`, `-Filtro-contenuti.log`) vengono riconosciuti automaticamente. Gli zip vengono letti in memoria, senza estrarre nulla su disco. Se lo stesso log è presente più volte, vale quello sciolto, poi quello più grande.

Con `--all-log` lo script legge solo `MDaemon-AAAA-MM-GG-all.log`, il log in cui MDaemon scrive tutto, e ignora gli altri log. Quel file può pesare centinaia di MB, perché contiene anche tutto il traffico IMAP: lo script ne mostra la dimensione e chiede conferma prima di leggerlo (`-y` o `--yes` salta la domanda; senza terminale e senza `-y` si ferma senza fare nulla). Viene letto a blocchi, anche da dentro uno zip, e dà gli stessi conteggi dei log separati.

### Configurazione (facoltativa)

- `aliases.txt`: `alias@dominio = principale@dominio`, una riga per alias.
- `mailboxes.txt`: `principale@dominio = C:\MDaemon\Users\dominio\cartella`, solo gli indirizzi principali. Serve quando il nome della cartella della casella non coincide con l'indirizzo.
- `log_names.txt`: nomi dei file di log per ogni lingua di MDaemon, una sezione per lingua. **Inglese** e **italiano** sono compilati; tedesco, francese, spagnolo, portoghese, russo, giapponese, coreano e cinese sono segnaposto vuoti, perché non è stata trovata una fonte affidabile per i loro nomi. Se i tuoi log non vengono trovati, togli il commento alla sezione della tua lingua e compilala con i nomi che vedi nella cartella `Logs` di MDaemon. Il file viene cercato accanto allo script o nella cartella dei log (`--log-names FILE` per usarne un altro); senza, vengono usati solo i nomi inglesi.

Vedi `aliases.example.txt` e `mailboxes.example.txt`. Quello che non è configurato viene ricavato dai log e segnalato nel foglio "Mailboxes and aliases", insieme alle contraddizioni tra configurazione e log.

### Metodo di conteggio

- Ogni messaggio entrato nel server (INBOUND) è contato una volta; i duplicati con lo stesso Message-ID e lo stesso mittente vengono uniti.
- **Inviate** (Sent): messaggi con mittente del dominio, divisi tra solo interni e con destinatari esterni (destinatari reali presi dal log SMTP in uscita). Risposte automatiche e notifiche del server sono tenute a parte.
- **Ricevute** (Received): consegne nelle caselle, una volta per messaggio e per casella, divise tra mittenti interni, esterni e notifiche del server.
- Con un intervallo di date i log sono analizzati insieme, quindi un messaggio a cavallo della mezzanotte resta uno solo.
- **Peso**: la dimensione di ogni messaggio è quella scritta nel log di instradamento quando il messaggio entra nel server. Un messaggio conta per il mittente e per ogni casella che lo riceve. I messaggi senza dimensione nei log (alcune notifiche del server) restano vuoti e non entrano in totali e medie; il loro numero è indicato a schermo e nel foglio "Sources and checks".

### Excel prodotto

Di default il file contiene solo il foglio **Summary by account** (riepilogo), con la media giornaliera e, per ogni account, il peso totale inviato e ricevuto (MB) e il peso medio per messaggio inviato e ricevuto (KB). I fogli **Message detail** e **Detail by account** hanno una colonna con il peso (KB), e **By day** i MB inviati e ricevuti per giorno. Gli altri fogli si aggiungono con opzioni brevi, combinabili tra loro:

| Opzione | Foglio |
|---------|--------|
| `-d` | By day (per giorno, solo con più giornate) |
| `-md` | Message detail (dettaglio messaggi) |
| `-ad` | Detail by account (dettaglio per account) |
| `-mb` | Mailboxes and aliases (caselle e alias) |
| `-s` | Sources and checks (fonti e verifiche) |
| `-a` | tutti i fogli |

Esempio: `python mail_counter.py --domain example.com --from 2026-01-10 --to 2026-01-15 -d -md`
