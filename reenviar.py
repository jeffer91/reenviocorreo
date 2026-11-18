import hashlib
import os
import re
import sqlite3
import time
from pathlib import Path

MARKER = "ReenvioCorreoOrigenV1"
FOLDER_TYPE = "http://schemas.microsoft.com/mapi/proptag/0x36010003"


def setting(db, name, prompt):
    row = db.execute("SELECT value FROM config WHERE name=?", (name,)).fetchone()
    if row:
        return row[0]
    value = input(prompt).strip().lower()
    if not re.fullmatch(r"[^@\s;:,<>]+@[^@\s;:,<>]+\.[^@\s;:,<>]+", value):
        raise ValueError("Direccion de correo invalida.")
    return value


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
    if original.Class != 43 or not original.Sent or original.UserProperties.Find(MARKER, True):
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


def monitor(db, session, account, target):
    store = account.DeliveryStore
    store_id = store.StoreID
    # Outbox and drafts contain messages which have not yet been submitted.
    excluded = {store.GetDefaultFolder(i).EntryID for i in (4, 16)}
    total = db.execute("SELECT COUNT(*) FROM progress WHERE status='enviado'").fetchone()[0]
    while True:
        if session.Offline:
            print("Outlook sin conexion. Esperando un minuto...", flush=True)
        else:
            for folder in folders(store.GetRootFolder(), excluded):
                print("Revisando", folder.FolderPath, flush=True)
                table = folder.GetTable()
                while not table.EndOfTable:
                    row = table.GetNextRow()
                    if not str(row.Item("MessageClass")).lower().startswith("ipm.note"):
                        continue
                    try:
                        submitted = forward(db, session, account, store_id, row.Item("EntryID"), target)
                    except Exception as exc:
                        import pywintypes
                        if not isinstance(exc, pywintypes.com_error):
                            raise
                        print(f"No se pudo preparar un mensaje en {folder.FolderPath}: {exc}", flush=True)
                        continue
                    if submitted:
                        total += 1
                        print(f"Entregados a Outlook para envio: {total}", flush=True)
                        time.sleep(3)
            print("Revision terminada. Proxima revision en un minuto.", flush=True)
        time.sleep(60)


def main():
    import msvcrt
    import win32com.client

    directory = Path(os.environ["LOCALAPPDATA"]) / "ReenvioCorreo"
    directory.mkdir(exist_ok=True)
    with (directory / "instancia.lock").open("a+b") as lock:
        lock.seek(0)
        try:
            msvcrt.locking(lock.fileno(), msvcrt.LK_NBLCK, 1)
        except OSError:
            raise RuntimeError("Ya hay otra instancia trabajando.") from None
        db = sqlite3.connect(directory / "estado.sqlite3")
        db.executescript("""
            CREATE TABLE IF NOT EXISTS config(name TEXT PRIMARY KEY, value TEXT NOT NULL);
            CREATE TABLE IF NOT EXISTS progress(id TEXT PRIMARY KEY, status TEXT NOT NULL);
        """)
        source = setting(db, "origen", "Correo de la cuenta de Outlook: ")
        target = setting(db, "destino", "Correo de destino: ")
        if source == target:
            raise ValueError("El origen y el destino deben ser distintos.")
        app = win32com.client.gencache.EnsureDispatch("Outlook.Application")
        session = app.Session
        account = next((a for a in session.Accounts if a.SmtpAddress.lower() == source), None)
        if account is None:
            raise RuntimeError("Abre Outlook clasico con la cuenta indicada y vuelve a iniciar.")
        if not db.execute("SELECT 1 FROM config LIMIT 1").fetchone():
            print(f"Se reenviaran el historial y los mensajes nuevos de {source} a {target}.")
            if input("Escribe INICIAR para comenzar: ").strip().upper() != "INICIAR":
                return
            with db:
                db.executemany("INSERT INTO config VALUES (?, ?)", [("origen", source), ("destino", target)])
        uncertain = db.execute("SELECT COUNT(*) FROM progress WHERE status='revisar'").fetchone()[0]
        if uncertain:
            print(f"Hay {uncertain} envios pendientes de verificar en Outlook/Gmail tras una interrupcion.")
        print(f"{source} -> {target}. Puedes minimizar la ventana. Ctrl+C para detener.")
        monitor(db, session, account, target)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("Detenido. El avance queda guardado.")
    except Exception as exc:
        print(f"Proceso detenido: {exc}", flush=True)
        raise SystemExit(1)
