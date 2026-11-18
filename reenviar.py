import hashlib
import os
import re
import sqlite3
import queue
import threading
import tkinter as tk
from contextlib import closing
from tkinter import ttk
from pathlib import Path

MARKER = "ReenvioCorreoOrigenV1"
FOLDER_TYPE = "http://schemas.microsoft.com/mapi/proptag/0x36010003"


def database(directory):
    db = sqlite3.connect(directory / "estado.sqlite3")
    db.executescript("""
        CREATE TABLE IF NOT EXISTS config(name TEXT PRIMARY KEY, value TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS progress(id TEXT PRIMARY KEY, status TEXT NOT NULL);
    """)
    return db


def received(original, account):
    if original.Class != 43 or not original.Sent or original.UserProperties.Find(MARKER, True):
        return False
    # Check the sender even for messages moved from Sent Items into an archive.
    sender = original.Sender
    if sender is None:
        return False
    owner = account.CurrentUser.AddressEntry
    if sender.ID and sender.ID == owner.ID:
        return False
    address = original.SenderEmailAddress or ""
    if original.SenderEmailType == "EX":
        if address.casefold() == (owner.Address or "").casefold():
            return False
        exchange_user = sender.GetExchangeUser()
        if exchange_user and exchange_user.PrimarySmtpAddress:
            address = exchange_user.PrimarySmtpAddress
    flags = original.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x0E070003")
    return bool(address) and not (flags & 0x20) and address.casefold() != account.SmtpAddress.casefold()


def folders(root, excluded):
    if root.EntryID in excluded or root.PropertyAccessor.GetProperty(FOLDER_TYPE) == 2:
        return
    if root.DefaultItemType == 0:
        yield root
    for child in root.Folders:
        yield from folders(child, excluded)


def forward(db, session, account, store_id, entry_id, target):
    key = hashlib.sha256((store_id + "/" + entry_id).encode()).hexdigest()
    if db.execute("SELECT 1 FROM progress WHERE id=?", (key,)).fetchone():
        return False
    original = session.GetItemFromID(entry_id, store_id)
    if not received(original, account):
        return False
    copy = original.Forward()
    copy.SendUsingAccount = account
    if copy.Recipients.Count:
        raise RuntimeError("El reenvio contiene destinatarios inesperados.")
    recipient = copy.Recipients.Add(target)
    if not recipient.Resolve():
        raise RuntimeError("Outlook no pudo resolver el destinatario.")
    copy.UserProperties.Add(MARKER, 1, False).Value = key
    # Persist before Send: a crash after submission must not trigger another send.
    with db:
        db.execute("INSERT INTO progress VALUES (?, 'revisar')", (key,))
    try:
        copy.Send()
    except Exception as exc:
        raise RuntimeError(f"Outlook rechazo el envio o no pudo confirmarlo: {exc}") from exc
    with db:
        db.execute("UPDATE progress SET status='enviado' WHERE id=?", (key,))
    return True


def monitor(db, session, account, target, stop, report):
    import pywintypes

    store = account.DeliveryStore
    store_id = store.StoreID
    excluded = {store.GetDefaultFolder(i).EntryID for i in (4, 5, 16)}
    total = 0
    while not stop.is_set():
        if session.Offline:
            report("status", "Outlook está sin conexión. Esperando…")
        else:
            errors = 0
            for folder in folders(store.GetRootFolder(), excluded):
                if stop.is_set():
                    return
                report("status", "Revisando " + folder.FolderPath)
                table = folder.GetTable()
                while not table.EndOfTable and not stop.is_set():
                    row = table.GetNextRow()
                    if not str(row.Item("MessageClass")).lower().startswith("ipm.note"):
                        continue
                    try:
                        submitted = forward(db, session, account, store_id, row.Item("EntryID"), target)
                    except pywintypes.com_error:
                        errors += 1
                        report("warning", f"Mensajes pendientes por error de lectura en esta revisión: {errors}")
                        continue
                    if submitted:
                        total += 1
                        report("count", str(total))
                        stop.wait(3)
            report("status", "Historial revisado. Próxima revisión en un minuto.")
        stop.wait(60)


def run(directory, source, target, stop, report):
    import msvcrt
    import win32com.client

    with (directory / "instancia.lock").open("a+b") as lock:
        lock.seek(0)
        try:
            msvcrt.locking(lock.fileno(), msvcrt.LK_NBLCK, 1)
        except OSError:
            raise RuntimeError("Otra instancia está activa. Detén el proceso anterior.") from None
        with closing(database(directory)) as db:
            app = win32com.client.gencache.EnsureDispatch("Outlook.Application")
            session = app.Session
            account = next((a for a in session.Accounts if a.SmtpAddress.casefold() == source), None)
            if account is None:
                raise RuntimeError("Abre Outlook clásico con la cuenta de origen indicada.")
            with db:
                db.executemany("INSERT OR IGNORE INTO config VALUES (?, ?)", [("origen", source), ("destino", target)])
            saved = dict(db.execute("SELECT name,value FROM config"))
            if (saved["origen"], saved["destino"]) != (source, target):
                raise RuntimeError("La configuración guardada corresponde a otra cuenta.")
            uncertain = db.execute("SELECT COUNT(*) FROM progress WHERE status='revisar'").fetchone()[0]
            if uncertain:
                report("warning", f"Envíos anteriores pendientes de comprobar: {uncertain}")
            monitor(db, session, account, target, stop, report)


def main():
    directory = Path(os.environ["LOCALAPPDATA"]) / "ReenvioCorreo"
    directory.mkdir(exist_ok=True)
    with closing(database(directory)) as db:
        saved = dict(db.execute("SELECT name,value FROM config"))
    root = tk.Tk()
    root.title("Reenvío de correo")
    panel = ttk.Frame(root, padding=24)
    panel.pack(fill="both", expand=True)
    panel.columnconfigure(0, weight=1)
    ttk.Label(panel, text="Correos recibidos", font=("Segoe UI", 18, "bold")).grid(sticky="w")
    ttk.Label(panel, text="Historial y mensajes nuevos · Leídos y sin leer").grid(sticky="w", pady=(0, 16))
    fields = []
    for key, label in [("origen", "Cuenta de Outlook"), ("destino", "Reenviar a")]:
        ttk.Label(panel, text=label).grid(sticky="w")
        entry = ttk.Entry(panel, width=52)
        entry.insert(0, saved.get(key, ""))
        entry.grid(sticky="ew", pady=(2, 10))
        if key in saved:
            entry.configure(state="readonly")
        fields.append(entry)
    status = tk.StringVar(value="Listo para iniciar")
    count = tk.StringVar(value="0")
    warning = tk.StringVar()
    ttk.Label(panel, textvariable=status, wraplength=410).grid(sticky="w", pady=(6, 4))
    ttk.Label(panel, text="Reenvíos entregados a Outlook en esta sesión").grid(sticky="w")
    ttk.Label(panel, textvariable=count, font=("Segoe UI", 18)).grid(sticky="w")
    ttk.Label(panel, textvariable=warning, wraplength=410).grid(sticky="w")
    actions = ttk.Frame(panel)
    actions.grid(sticky="w", pady=(10, 0))
    events = queue.Queue()
    stop = threading.Event()
    active = False
    closing_window = False

    def worker(source, target):
        try:
            import pythoncom
            pythoncom.CoInitialize()
            try:
                run(directory, source, target, stop, lambda k, v: events.put((k, v)))
            finally:
                pythoncom.CoUninitialize()
            events.put(("done", "Detenido. Avance guardado."))
        except Exception as exc:
            events.put(("done", str(exc)))

    def start():
        nonlocal active
        source, target = [entry.get().strip().casefold() for entry in fields]
        if source == target or not all(re.fullmatch(r"[^@\s;:,<>]+@[^@\s;:,<>]+\.[^@\s;:,<>]+", addr) for addr in (source, target)):
            status.set("Introduce dos direcciones de correo válidas y distintas.")
            return
        stop.clear()
        active = True
        for entry in fields:
            entry.configure(state="readonly")
        start_button.configure(state="disabled")
        stop_button.configure(state="normal")
        status.set("Conectando con Outlook…")
        warning.set("")
        count.set("0")
        threading.Thread(target=worker, args=(source, target)).start()

    def halt():
        stop.set()
        status.set("Deteniendo después del mensaje en curso…")
        stop_button.configure(state="disabled")

    def close():
        nonlocal closing_window
        if active:
            closing_window = True
            halt()
        else:
            root.destroy()

    def refresh():
        nonlocal active
        while not events.empty():
            kind, value = events.get_nowait()
            if kind == "done":
                active = False
                if closing_window:
                    root.destroy()
                    return
                with closing(database(directory)) as db:
                    configured = dict(db.execute("SELECT name,value FROM config"))
                for key, entry in zip(("origen", "destino"), fields):
                    entry.configure(state="readonly" if key in configured else "normal")
                start_button.configure(state="normal")
                stop_button.configure(state="disabled")
                status.set(value)
            elif kind != "status" or not stop.is_set():
                {"status": status, "warning": warning, "count": count}[kind].set(value)
        root.after(200, refresh)

    start_button = ttk.Button(actions, text="Iniciar", command=start)
    start_button.pack(side="left")
    stop_button = ttk.Button(actions, text="Detener", command=halt, state="disabled")
    stop_button.pack(side="left", padx=8)
    root.protocol("WM_DELETE_WINDOW", close)
    root.after(200, refresh)
    root.mainloop()


if __name__ == "__main__":
    main()
