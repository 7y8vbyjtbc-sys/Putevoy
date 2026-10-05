#!/usr/bin/env python3
"""
Консольная версия генератора путевых листов.

  python putevoy.py ПЛ_за_прошлый_месяц.xlsm
  python putevoy.py ПЛ.xlsm --liters 300 --norm 7.5 --office "г. Мурманск, Карла-Либкнехта 28" --yes
  python putevoy.py --import-shops Список_ММ.xlsx      # обновить справочник магазинов
"""
import argparse
import os
import sys

import putevoy_core as C

HERE = os.path.dirname(os.path.abspath(__file__))


def ask(prompt, default=None):
    suffix = f" [{default}]" if default not in (None, "") else ""
    ans = input(f"{prompt}{suffix}: ").strip()
    return ans or ("" if default is None else str(default))


def ask_number(prompt, default=None):
    while True:
        try:
            return C.parse_decimal(ask(prompt, default), prompt)
        except C.UserError as e:
            print(f"  {e}")


def ask_changes(consts):
    print("\nПостоянные данные из прошлого листа:")
    for cell, label, _ in C.FIELDS:
        print(f"  {label}: {C.format_field(consts.get(cell))}")
    if ask("Что-то из этого изменилось? (д/н)", "н").lower() not in ("д", "да", "y", "yes"):
        return {}
    print("Введите новое значение или нажмите Enter, чтобы оставить как есть.")
    changes = {}
    for cell, label, kind in C.FIELDS:
        while True:
            new = input(f"  {label} [{C.format_field(consts.get(cell))}]: ").strip()
            if not new:
                break
            try:
                changes[cell] = C.parse_field(new, kind)
                break
            except C.UserError as e:
                print(f"  {e}")
    return changes


def main():
    ap = argparse.ArgumentParser(description="Генератор путевого листа на следующий месяц")
    ap.add_argument("prev", nargs="?", help="путевой лист за прошлый месяц (.xlsm)")
    ap.add_argument("--liters", help="литров выдано за месяц")
    ap.add_argument("--norm", help="норма расхода на 100 км")
    ap.add_argument("--office", help="адрес выезда (по умолчанию — из прошлого листа)")
    ap.add_argument("--shops", default=os.path.join(HERE, "shops.json"), help="справочник магазинов")
    ap.add_argument("--out", help="куда сохранить (файл или папка)")
    ap.add_argument("--seed", type=int, help="зерно случайности для повторяемого результата")
    ap.add_argument("--weekends", action="store_true", help="разрешить поездки в выходные")
    ap.add_argument("--yes", action="store_true", help="не задавать вопросов о постоянных данных и адресе")
    ap.add_argument("--import-shops", metavar="XLSX", help="обновить справочник из списка магазинов")
    args = ap.parse_args()

    db = C.load_db(args.shops)
    try:
        if args.import_shops:
            rep = C.import_shops(args.import_shops, db)
            C.save_db(db, args.shops)
            print(f"Справочник обновлён: {rep['total']} магазинов "
                  f"(новых {len(rep['added'])}, удалено {len(rep['removed'])}, сменили адрес {len(rep['changed'])})")
            return
        if not args.prev:
            ap.error("укажите путевой лист за прошлый месяц")

        info = C.read_prev(args.prev)
        head, sub = C.describe_prev(info)
        print(f"\nПрошлый лист: {head}\n{sub}")
        if info["warning"]:
            print("! " + info["warning"])

        office = args.office or info["office"]
        if not args.yes and not args.office:
            office = ask("\nАдрес выезда (Enter — без изменений)", office)
        changes = {} if args.yes else ask_changes(info["consts"])
        liters = C.parse_decimal(args.liters, "Литры") if args.liters else ask_number("\nСколько литров выдано за месяц")
        norm = C.parse_decimal(args.norm, "Норма") if args.norm else ask_number("Норма расхода на 100 км", info["norm"])

        res = C.generate(info, db, office, liters, norm, args.weekends, args.seed)
        name = C.output_name(info, res, changes)
        if args.out:
            out = os.path.join(args.out, name) if os.path.isdir(args.out) else args.out
        else:
            out = os.path.join(os.path.dirname(os.path.abspath(args.prev)), name)
        C.write_result(info, res, out, changes)
        print(f"\nАдрес выезда: {office} (населённый пункт: {res['office_town']}, магазинов доступно: {res['points']})\n")
        print(C.summary_text(res))
        print(f"\nГотово: {out}")
    except C.UserError as e:
        sys.exit(f"Ошибка: {e}")


if __name__ == "__main__":
    main()
