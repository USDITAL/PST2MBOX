#!/usr/bin/env python3
"""Diagnóstico completo de estructura PST"""
import sys
import pypff

def dump_item(item, depth=0):
    indent = "  " * depth
    try:
        n_rs = item.number_of_record_sets
        n_items = item.number_of_sub_items
        print(f"{indent}ITEM id={item.identifier} record_sets={n_rs} sub_items={n_items}")
        
        # Intentar leer propiedades MAPI clave
        for rs_i in range(min(n_rs, 1)):
            rs = item.get_record_set(rs_i)
            for e_i in range(rs.number_of_entries):
                entry = rs.get_entry(e_i)
                etype = entry.entry_type
                # Propiedades interesantes
                if etype in (0x0037, 0x0C1A, 0x0E04, 0x1000, 0x1013):
                    try:
                        val = entry.data_as_string[:60]
                        names = {0x0037:'Subject', 0x0C1A:'SenderName',
                                 0x0E04:'To', 0x1000:'Body', 0x1013:'HTMLBody'}
                        print(f"{indent}  [{names.get(etype, hex(etype))}] = {val!r}")
                    except Exception:
                        print(f"{indent}  [{hex(etype)}] type={entry.value_type} data_len={len(entry.data or b'')}")
    except Exception as e:
        print(f"{indent}  ERROR leyendo item: {e}")

def explore_folder(folder, depth=0):
    indent = "  " * depth
    try:
        name = folder.name or "?"
    except:
        name = "?"
    
    try: n_msg   = folder.number_of_sub_messages
    except: n_msg = "ERR"
    try: n_items = folder.number_of_sub_items
    except: n_items = "ERR"
    try: n_sub   = folder.number_of_sub_folders
    except: n_sub = "ERR"
    try: n_entries = folder.number_of_entries
    except: n_entries = "ERR"

    print(f"{indent}FOLDER [{name}]")
    print(f"{indent}  sub_messages={n_msg}  sub_items={n_items}  sub_folders={n_sub}  entries={n_entries}")

    # Mostrar primeros 3 sub_items para ver qué tipo son
    if isinstance(n_items, int) and n_items > 0:
        print(f"{indent}  → Inspeccionando sub_items (max 3):")
        for i in range(min(n_items, 3)):
            try:
                item = folder.get_sub_item(i)
                dump_item(item, depth + 2)
            except Exception as e:
                print(f"{indent}    sub_item[{i}] ERROR: {e}")

    # Mostrar primeros 3 sub_messages
    if isinstance(n_msg, int) and n_msg > 0:
        print(f"{indent}  → Inspeccionando sub_messages (max 3):")
        for i in range(min(n_msg, 3)):
            try:
                msg = folder.get_sub_message(i)
                print(f"{indent}    msg[{i}]: subject={msg.subject!r}")
            except Exception as e:
                print(f"{indent}    sub_message[{i}] ERROR: {e}")

    # Recursión en subcarpetas
    if isinstance(n_sub, int) and n_sub > 0:
        for i in range(n_sub):
            try:
                explore_folder(folder.get_sub_folder(i), depth + 1)
            except Exception as e:
                print(f"{indent}  subfolder[{i}] ERROR: {e}")

if len(sys.argv) < 2:
    print("Uso: python3 diagnose_pst.py archivo.pst")
    sys.exit(1)

pff = pypff.file()
pff.open(sys.argv[1])
root = pff.get_root_folder()

print(f"=== ESTRUCTURA PST: {sys.argv[1]} ===\n")
print(f"ROOT: sub_folders={root.number_of_sub_folders} sub_messages={root.number_of_sub_messages} sub_items={root.number_of_sub_items}\n")

for i in range(root.number_of_sub_folders):
    explore_folder(root.get_sub_folder(i), depth=0)

pff.close()
print("\n=== FIN DIAGNÓSTICO ===")