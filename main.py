from table_lookup import TableLookup

# главная таблица — только пишем в неё, не меняем
lookup = TableLookup(
    path="tests/manual tests/Заявка_главная.xlsx",
    key_column="Договор контрагента",
    sheet_name=None,
    output_column="Outer",  # None → первый безымянный столбик после последнего
    # output_column="Комментарий",  # или явно
)

# источник 1: реестр выполненных
lookup.addSourceTable(
    path="tests/manual tests/Реестр выполненных заявок.xlsx",
    sheet_name="ФЛ",
    key_column="Лицевой счет",
    value_column="ДАТА Установки ПУ",
    template="дата установки: {value}",
)

# источник 2: невыполненные — физ. лица
lookup.addSourceTable(
    path="tests/manual tests/Реестр невыполненных заявок.xlsx",
    sheet_name="ФЛ",
    key_column="ЛС",
    value_column="Источник заявки",
    template="Источник заявки: {value}",
)

lookup.run()
