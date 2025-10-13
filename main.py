from docx import Document


def _replace_in_paragraph_runs(paragraph, placeholder, value):
    """
    Аккуратно заменяет placeholder на value в paragraph, не теряя форматирование абзаца.
    Сохраняем форматирование первого run'а, в котором начинается найденный placeholder.
    """
    # Соберем склейку всех run'ов, чтобы найти индексы в едином тексте
    runs = paragraph.runs
    if not runs:
        return

    full_text = "".join(run.text for run in runs)
    start = full_text.find(placeholder)
    while start != -1:
        end = start + len(placeholder)

        # Найдем соответствие "позиция в полном тексте -> (run_index, offset)"
        run_positions = []
        pos = 0
        for i, r in enumerate(runs):
            text = r.text
            run_positions.append((i, pos, pos + len(text)))  # (индекс run, начало, конец в full_text)
            pos += len(text)

        # Определим run, где начинается и заканчивается placeholder
        # start находится в промежутке [run_start, run_end)
        def locate(idx):
            for (ri, s, e) in run_positions:
                if s <= idx < e:
                    return ri, idx - s
            # если idx == длине текста, привяжем к концу последнего run
            return len(runs) - 1, len(runs[-1].text)

        start_ri, start_off = locate(start)
        end_ri, end_off = locate(end - 1)  # последняя буква плейсхолдера

        # Три части: before (в run start), middle (полностью покрытые run'ы), after (в run end)
        before = runs[start_ri].text[:start_off]
        after = runs[end_ri].text[end_off + 1:]  # +1 т.к. end_off указывает на последнюю букву плейсхолдера

        # В первый run пишем before + value + (если плейсхолдер полностью в одном run) after
        runs[start_ri].text = before + value + (after if start_ri == end_ri else "")

        # Все run'ы между началом и концом чистим
        for i in range(start_ri + 1, end_ri):
            runs[i].text = ""

        if end_ri != start_ri:
            # В конечном run оставляем только хвост после плейсхолдера
            runs[end_ri].text = after

        # Пересоберем full_text после правки и ищем следующий вхождение
        full_text = "".join(run.text for run in runs)
        start = full_text.find(placeholder, start + len(value))


def _iter_table_cells(table):
    for row in table.rows:
        for cell in row.cells:
            yield cell
            for t in cell.tables:
                yield from _iter_table_cells(t)  # вложенные таблицы


def _replace_in_container(container, placeholder, value):
    # Абзацы
    for p in container.paragraphs:
        _replace_in_paragraph_runs(p, placeholder, value)
    # Таблицы (включая вложенные)
    for t in getattr(container, "tables", []):
        for cell in _iter_table_cells(t):
            for p in cell.paragraphs:
                _replace_in_paragraph_runs(p, placeholder, value)


def fill_template(template_path, output_path, context):
    """
    Заменяет плейсхолдеры {{ VAR }} в теле документа, таблицах и колонтитулах.
    """
    doc = Document(template_path)

    def do_all_containers(placeholder, value):
        # Тело документа
        _replace_in_container(doc, placeholder, value)
        # Колонтитулы всех секций
        for section in doc.sections:
            _replace_in_container(section.header, placeholder, value)
            _replace_in_container(section.footer, placeholder, value)

    for key, val in context.items():
        placeholder = f"{{{{ {key} }}}}"  # {{ ORG_NAME }}
        do_all_containers(placeholder, val)

    doc.save(output_path)
    print(f"✅ Документ сохранён: {output_path}")


if __name__ == "__main__":
    # ==== 🔧 ДАННЫЕ ДЛЯ ЗАПОЛНЕНИЯ ====
    context = {
        # --- Основная информация об экспертизе ---
        "EXP_NUMBER": "171",
        "EXP_DATE": "13.10.2025",  # дата ЭПБ
        "ORDER_DATE": "13.09.2025",  # дата приказа
        "EXP_YEAR": "2025",

        # --- Организация ---
        "ORG_NAME": "Башнефть-Добыча",
        "ORG_LEGAL_FORM": "Общество с ограниченной ответственностью",
        "ORG_LEGAL_FORM_SHORT": "ООО",
        "ORG_INN": "0277106840",
        "ORG_OGRN": "1090280032699",
        "ORG_ADDRESS": "450052, Республика Башкортостан, г. Уфа, ул. Карла Маркса, д. 30/1",
        "ORG_PHONE": "+7 (347) 261-61-61",
        "ORG_EMAIL": "info_bn@bashneft.ru",

        # --- Руководитель ---
        "DECLARATION_APPROVER": "Генеральный директор",
        "DECLARATION_APPROVER_FIO": "Нонява Сергей Александрович",
        "DECLARATION_DATE": "10.10.2025",

        # --- ОПО ---
        "OPO_NAME": "Система промысловых трубопроводов Белебеевского месторождения",
        "OPO_ADDRESS": "Республика Башкортостан, Белебеевский  р-н, Ермекеевский  р-н, Бижбулякский  р-н",
        "OPO_CLASS": "II",
        "REG_NUMBER": "А41-05127-0365",
        "OPO_SUPERIOR_ORG": "ПАО «НК «Роснефть»",
        "OPO_DESCRIPTION": (
            "Система промысловых трубопроводов Белебеевского месторождения ООО «Башнефть-Добыча» "
            "включает в состав трубопроводы для сбора и транспортирования продукции скважин "
            "Белебеевского нефтяного место-рождения до замерных установок, от замерных установок до "
            "УПОСН «Белебей» и далее до ППН «Чегодаево»."
        ),
        "OPO_COMPONENTS": "Система промысловых трубопроводов Белебеевского месторождения",
        "OPO_SANZONE": "25 м",

        # --- Лицензии и резервы ---
        "LICENSE_NUMBER": "ВХ-00-015657 от 09.10.2015",
        "ORDER_MAT_RESERVES": (
            "№ 0445 от 06.04.2020 г. "
            "«О создании резерва материальных ресурсов для локализации и ликвидации последствий аварий "
            "на опасных производственных объектах»"
        ),
        "ORDER_FIN_RESERVES": (
            "№ 0112 от 10 февраля 2025 г. Инструкция ООО «Башнефть-Добыча» "
            "«Создание и использование финансового резерва для ликвидации чрезвычайных ситуаций "
            "природного и техногенного характера» № ПЗ-11.04И-001112 ЮЛ-305, версия 2"
        ),
        "ORDER_NASF": "№ 0127 от 02.02.2022 г.",

        # --- НАСФ, ПАСФ, ПЧ ---
        "NASF_CERTIFICATE": "№16423 от 22.04.2025 г., рег.№16/2-2-477",  # номер свидетельства НАСФ
        "PASF_NAME": "ФГАУ «АСФ «СВПФВЧ»",  # наименование ПАСФ
        "PASF_CERTIFICATE": "№13307 от 01.07.2022 г., рег.№8-177",  # номер свидетельства
        "PASF_CONTRACT": "№БНД/У/8/1030/21/ОПБ от 28.10.2021 г.",  # договор с ПАСФ
        "FIRE_DEPARTMENT_NAME": "ООО «РН-Пожарная безопасность»",  # наименование ПЧ
        "FIRE_DEPARTMENT_CONTRACT": "БНД/У/8/1444/23/ОПБ//5702623/0774Д от 18.12.2023 г.",  # договор с ПЧ

        # --- Объемы документов ---
        "DECLARATION_VOLUME": "59",
        "RPZ_VOLUME": "132",
        "IFL_VOLUME": "9",

        # --- Документы при экспертизе ---
        "DOCS_SUBMITTED": (
            "Технологический регламент ОПО «Система промысловых трубопроводов Белебеевского месторождения», План мероприятий по локализации и ликвидации последствий аварий на опасных производственных объектах систем "
            "промысловых трубопроводов Ишимбайского региона ДНГ ИЦТОиРТ УЭТ ООО «Башнефть-добыча»."
            "№ П3-05 ПМ-0359 ЮЛ-305 "
        ),

        # --- Разработчик декларации ---
        "DECLARATION_DEVELOPER": "ООО «Экопромпроект»",
        "DECLARATION_DEVELOPER_ADDRESS": "420127, Республика Татарстан, г. Казань, ул. Короленко, д. 120, офис 1",
        "DECLARATION_DEVELOPER_PHONE": "+7-917-286-81-25",
        "DECLARATION_DEVELOPER_EMAIL": "ildarkalimullin@gmail.com",
    }

    from pathlib import Path

    # Пакетные задания: тот же набор плейсхолдеров → разные шаблоны/результаты
    jobs = [
        ("template2.docx", "pismo.docx"),  # шаблон письма → итог pismo.docx
        ("template.docx", "epb.docx"),  # шаблон ЭПБ → итог epb.docx
    ]

    for tpl, out in jobs:
        if not Path(tpl).exists():
            print(f"⚠️ Шаблон не найден: {tpl}. Пропускаю.")
            continue
        fill_template(tpl, out, context)
        print(f"✅ Заполнено: {tpl} → {out}")
