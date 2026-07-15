#Requires AutoHotkey v2.0
#SingleInstance Force

; ============================================================
;  Parser v2.1 — универсальный парсер коммерческих предложений
;  Работает через буфер обмена (только чтение/запись A_Clipboard,
;  без автокопирования и автовставки — вставляет пользователь сам)
;
;  Ctrl+Shift+X — КП -> Таблица
;  Ctrl+Shift+C — Таблица -> КП
;  Ctrl+Shift+D — Диагностика (окно со скроллом)
;  Ctrl+Shift+N — Краткий список "Наименование - Количество"
;
;  Изменения v2.1 (правки после ревью v2.0):
;   - NUM объявлена как super-global и явно декларируется global
;     внутри всех функций, где используется (в v2.0 без этого
;     скрипт падал с ошибкой "variable has not been assigned").
;   - Убран опасный fallback в ParsePrice(), который мог принять
;     число количества за цену, если "руб" в строке не найдено.
;   - Округление сумм до целого числа (FormatNum больше не
;     показывает копейки).
;   - Добавлены коды ошибок (E001-E007) для диагностики.
;   - Диагностика (Ctrl+Shift+D) теперь в отдельном окне со
;     скроллом вместо MsgBox, и умеет определять оба направления
;     (КП или таблица во входных данных).
;   - Обработка пустого/нетекстового буфера обмена через ClipWait.
;   - Нормализация неразрывного пробела (Word/Excel copy-paste).
;   - Строки таблицы с некорректным числом колонок теперь не
;     пропускаются молча, а попадают в диагностику как ошибка.
;   - Исправлен перепутанный порядок аргументов TrayTip().
; ============================================================

; "Константа" — количество: до 3 цифр, опционально группы по 3 через
; пробел (тысячи), опционально дробная часть через запятую/точку.
; Объявлена super-global, но каждая функция ниже дополнительно
; декларирует "global NUM" на всякий случай (см. комментарий выше).
global NUM := "\d{1,3}(?:[ ]\d{3})*(?:[.,]\d+)?"

TrayTip("Скрипт запущен и готов к работе.", "Parser v2.1")

; ------------------------------------------------------------
; Hotkeys
; ------------------------------------------------------------

^+x:: KPToTable()
^+c:: TableToKP()
^+d:: DebugMode()
^+n:: ShortList()

; ------------------------------------------------------------
; Работа с буфером обмена
; ------------------------------------------------------------

; Возвращает текст из буфера обмена или "" если буфер пуст /
; не содержит текста (например, скопировано изображение).
GetClipboardText() {
    try {
        if !ClipWait(0.5, 0)
            return ""
        return A_Clipboard
    } catch {
        return ""
    }
}

; ------------------------------------------------------------
; Normalize() — нормализация исходного текста
; ------------------------------------------------------------

Normalize(text) {
    ; неразрывный пробел (частый гость при копировании из Word/Excel/1С) -> обычный
    text := StrReplace(text, Chr(0xA0), " ")

    ; унификация тире/дефисов
    text := StrReplace(text, "—", "-")
    text := StrReplace(text, "–", "-")
    text := StrReplace(text, "−", "-")

    ; унификация "р/шт", "руб/шт", "руб / шт", "р / шт" -> "руб/шт"
    text := RegExReplace(text, "i)\bр(?:уб)?\s*/\s*шт\b", "руб/шт")

    ; убираем только повторяющиеся ПРОБЕЛЫ (не трогаем табы —
    ; они разделяют колонки таблицы, и не трогаем одиночные пробелы,
    ; чтобы "А4 200гр" не превратилось в "А4200гр")
    text := RegExReplace(text, "[ ]{2,}", " ")

    return text
}

; ------------------------------------------------------------
; Tokenizer() — разбивает строку на "название" и "расчётную часть"
; Якорь разделения — признаки цены: "руб"/"р" или "шт" (см. LooksLikeCalc)
; ------------------------------------------------------------

Tokenizer(line) {
    global NUM
    line := Trim(line)
    if (line = "")
        return {name: "", calc: ""}

    ; --- Проход 1: перебираем все " - " с конца, берём первый (с конца),
    ;     чей "хвост" похож на расчётную часть по якорю руб/шт ---
    positions := []
    searchPos := 1
    while (foundPos := InStr(line, " - ", , searchPos)) {
        positions.Push(foundPos)
        searchPos := foundPos + 1
    }
    loop positions.Length {
        idx := positions.Length - A_Index + 1
        dashPos := positions[idx]
        namePart := SubStr(line, 1, dashPos - 1)
        calcPart := SubStr(line, dashPos + 3)
        if LooksLikeCalc(calcPart)
            return {name: Trim(namePart), calc: Trim(calcPart)}
    }

    ; --- Проход 2: ищем начало расчётной части без дефиса ---
    ; (например "Ручка 85 руб/шт х 50шт = 4250 руб")
    pattern := "i)(?:(?:" . NUM . ")\s*шт\s*(?:х|x)\s*)?(" . NUM . ")\s*(?:руб|р)(?:\s*/\s*шт)?\b.*$"
    if RegExMatch(line, pattern, &m) {
        startPos := m.Pos(0)
        namePart := SubStr(line, 1, startPos - 1)
        calcPart := SubStr(line, startPos)
        return {name: Trim(namePart, " -,`t"), calc: Trim(calcPart)}
    }

    ; --- Проход 3: строка целиком либо расчёт, либо название ---
    if LooksLikeCalc(line)
        return {name: "", calc: line}

    return {name: line, calc: ""}
}

; строка "похожа на расчётную часть", если содержит признаки цены/количества
LooksLikeCalc(s) {
    return RegExMatch(s, "i)(руб|р\s*/\s*шт|\bшт\b)") ? true : false
}

; ------------------------------------------------------------
; Parser — отдельные функции разбора
; ------------------------------------------------------------

ParseName(nameStr) {
    return Trim(nameStr, " -,`t")
}

ParseQty(calc) {
    global NUM
    if RegExMatch(calc, "i)(" . NUM . ")\s*шт", &m)
        return CleanNum(m[1])
    return 1
}

ParsePrice(calc) {
    global NUM
    ; 1) цена за штуку "NUM руб/шт" / "NUM р/шт"
    if RegExMatch(calc, "i)(" . NUM . ")\s*(?:руб|р)\s*/\s*шт", &m)
        return CleanNum(m[1])

    ; 2) если есть "=", цену ищем в части ДО знака "="
    if InStr(calc, "=") {
        beforeEq := SubStr(calc, 1, InStr(calc, "=") - 1)
        if RegExMatch(beforeEq, "i)(" . NUM . ")\s*(?:руб|р)\b", &m2)
            return CleanNum(m2[1])
    }

    ; 3) одиночное "NUM руб" (когда нет разделения на цену/итог)
    if RegExMatch(calc, "i)(" . NUM . ")\s*(?:руб|р)\b", &m3)
        return CleanNum(m3[1])

    ; НЕТ запасного варианта "просто число" — если "руб"/"р" нигде не
    ; найдено, значит цены в строке действительно нет. В v2.0 здесь был
    ; fallback, который мог принять количество ("5шт") за цену и молча
    ; насчитать неверную сумму. Лучше явная ошибка E004, чем выдуманная цена.
    return 0
}

ParseTotal(calc, qty, price) {
    global NUM
    if RegExMatch(calc, "=\s*(" . NUM . ")", &m)
        return CleanNum(m[1])
    return qty * price
}

; собирает одну позицию из строки (название + расчёт уже объединены через " - ")
ParseKPLine(line) {
    line := Trim(line)
    if (line = "")
        return {ok: false, code: "E001", error: "Пустая строка", raw: line}

    tok := Tokenizer(line)
    if (tok.calc = "")
        return {ok: false, code: "E002", error: "Не найдена расчётная часть (цена/шт)", raw: line}

    name := ParseName(tok.name)
    if (name = "")
        return {ok: false, code: "E003", error: "Не найдено название", raw: line}

    qty := ParseQty(tok.calc)
    price := ParsePrice(tok.calc)
    if (price = 0)
        return {ok: false, code: "E004", error: "Не удалось распознать цену", raw: line}

    total := ParseTotal(tok.calc, qty, price)
    return {ok: true, name: name, qty: qty, price: price, total: total, raw: line}
}

; ------------------------------------------------------------
; Сборка позиций из текста КП (учитывает одно- и двухстрочный формат)
; ------------------------------------------------------------

BuildItemsFromKP(text) {
    text := Normalize(text)
    items := []

    rawLines := StrSplit(text, "`n", "`r")
    lines := []
    for _, l in rawLines {
        if (Trim(l) != "")
            lines.Push(l)
    }

    i := 1
    n := lines.Length
    while (i <= n) {
        line := lines[i]
        tok := Tokenizer(line)

        if (tok.name != "" && tok.calc != "") {
            ; однострочная позиция
            items.Push(ParseKPLine(line))
            i += 1
        } else if (tok.name != "" && tok.calc = "") {
            ; возможно двухстрочный формат: название + расчёт на след. строке
            matched := false
            if (i < n) {
                tok2 := Tokenizer(lines[i + 1])
                if (tok2.name = "" && tok2.calc != "") {
                    combined := tok.name . " - " . tok2.calc
                    items.Push(ParseKPLine(combined))
                    i += 2
                    matched := true
                }
            }
            if !matched {
                items.Push({ok: false, code: "E002", error: "Не найдена расчётная часть (цена/шт)", raw: line})
                i += 1
            }
        } else {
            items.Push({ok: false, code: "E005", error: "Не удалось разобрать строку (формат не распознан)", raw: line})
            i += 1
        }
    }
    return items
}

; ------------------------------------------------------------
; Сборка позиций из таблицы (Tab-разделённые колонки — так нужно
; для совместимости с Excel / БизнесПак при вставке результата)
; ------------------------------------------------------------

BuildItemsFromTable(text) {
    items := []
    lines := StrSplit(text, "`n", "`r")
    for _, l in lines {
        l := Trim(l)
        if (l = "")
            continue
        if RegExMatch(l, "i)^Наименование\b")
            continue ; пропускаем строку заголовка

        cols := StrSplit(l, "`t")
        if (cols.Length < 3) {
            items.Push({ok: false, code: "E006", error: "Менее 3 колонок (ожидается Tab-разделение: Наименование/Количество/Цена)", raw: l})
            continue
        }

        name := Trim(cols[1])
        qty := CleanNum(cols[2])
        price := CleanNum(cols[3])
        total := (cols.Length >= 4 && Trim(cols[4]) != "") ? CleanNum(cols[4]) : qty * price

        items.Push({ok: true, name: name, qty: qty, price: price, total: total})
    }
    return items
}

; ------------------------------------------------------------
; КП -> Таблица (Ctrl+Shift+X)
; ------------------------------------------------------------

KPToTable() {
    text := GetClipboardText()
    if (text = "") {
        MsgBox("Буфер обмена пуст или не содержит текста.", "Parser v2.1", "Icon!")
        return
    }

    items := BuildItemsFromKP(text)
    out := ""
    okCount := 0
    for _, it in items {
        if (it.ok) {
            out .= it.name . "`t" . FormatNum(it.qty) . "`t" . FormatNum(it.price) . "`t" . FormatNum(it.total) . "`n"
            okCount += 1
        }
    }

    if (okCount = 0) {
        MsgBox("Не удалось распознать ни одной строки.`nИспользуйте Ctrl+Shift+D для диагностики.", "Parser v2.1", "Icon!")
        return
    }

    A_Clipboard := RTrim(out, "`n")
    ToolTip("Готово: " . okCount . " из " . items.Length . " строк распознано. Результат в буфере обмена.")
    SetTimer(() => ToolTip(), -2500)
}

; ------------------------------------------------------------
; Таблица -> КП (Ctrl+Shift+C)
; ------------------------------------------------------------

TableToKP() {
    text := GetClipboardText()
    if (text = "") {
        MsgBox("Буфер обмена пуст или не содержит текста.", "Parser v2.1", "Icon!")
        return
    }

    rawItems := BuildItemsFromTable(text)
    items := []
    errCount := 0
    for _, it in rawItems {
        if (it.ok)
            items.Push(it)
        else
            errCount += 1
    }

    if (items.Length = 0) {
        MsgBox("Не удалось распознать ни одной строки таблицы.`nОжидается формат: Наименование [Tab] Количество [Tab] Цена [Tab] Стоимость (необязательно).`nИспользуйте Ctrl+Shift+D для диагностики.", "Parser v2.1", "Icon!")
        return
    }

    out := ""
    sum := 0
    for _, it in items {
        out .= it.name . " - " . FormatNum(it.qty) . "шт х " . FormatNum(it.price) . " руб/шт = " . FormatNum(it.total) . " руб`n"
        sum += it.total
    }
    out .= "`nИтого к оплате - " . FormatNum(sum) . " руб`n`nОплата наличными / счет на Юр. лицо"

    A_Clipboard := out
    msg := "Готово: КП сформировано (" . items.Length . " позиций)."
    if (errCount > 0)
        msg .= " Пропущено строк с ошибками: " . errCount . " (см. Ctrl+Shift+D)."
    ToolTip(msg)
    SetTimer(() => ToolTip(), -3000)
}

; ------------------------------------------------------------
; Краткий список (Ctrl+Shift+N)
; ------------------------------------------------------------

ShortList() {
    text := GetClipboardText()
    if (text = "") {
        MsgBox("Буфер обмена пуст или не содержит текста.", "Parser v2.1", "Icon!")
        return
    }

    items := []
    if (InStr(text, "`t") && RegExMatch(text, "i)Наименование")) {
        for _, it in BuildItemsFromTable(text) {
            if (it.ok)
                items.Push(it)
        }
    } else {
        for _, it in BuildItemsFromKP(text) {
            if (it.ok)
                items.Push(it)
        }
    }

    if (items.Length = 0) {
        MsgBox("Не удалось распознать ни одной позиции.", "Parser v2.1", "Icon!")
        return
    }

    out := ""
    for _, it in items
        out .= it.name . " - " . FormatNum(it.qty) . " шт`n"

    A_Clipboard := RTrim(out, "`n")
    ToolTip("Готово: краткий список (" . items.Length . " позиций) в буфере обмена.")
    SetTimer(() => ToolTip(), -2500)
}

; ------------------------------------------------------------
; Диагностика (Ctrl+Shift+D) — окно со скроллом, оба направления
; ------------------------------------------------------------

DebugMode() {
    text := GetClipboardText()
    if (text = "") {
        MsgBox("Буфер обмена пуст или не содержит текста.", "Parser v2.1", "Icon!")
        return
    }

    isTable := InStr(text, "`t") && RegExMatch(text, "i)Наименование")
    items := isTable ? BuildItemsFromTable(text) : BuildItemsFromKP(text)

    okList := []
    failList := []
    for _, it in items {
        if (it.ok)
            okList.Push(it)
        else
            failList.Push(it)
    }

    report := "Режим входных данных: " . (isTable ? "Таблица (Tab-разделённая)" : "КП (текст)") . "`r`n"
    report .= "Всего строк/позиций: " . items.Length . "`r`n"
    report .= "Распознано успешно: " . okList.Length . "`r`n"
    report .= "Не распознано: " . failList.Length . "`r`n"

    if (failList.Length > 0) {
        report .= "`r`n--- Проблемные строки ---`r`n"
        for _, f in failList
            report .= "`r`n[" . f.code . "] " . f.error . "`r`nСтрока: " . f.raw . "`r`n"
    } else {
        report .= "`r`nВсе строки распознаны успешно."
    }

    ShowDebugGui(report)
}

; Отдельное окно с прокруткой вместо MsgBox — длинные КП с кучей
; проблемных строк в MsgBox просто не помещались нормально.
ShowDebugGui(report) {
    dbgGui := Gui("+Resize", "Parser v2.1 — Диагностика")
    dbgGui.SetFont("s10", "Consolas")
    edit := dbgGui.AddEdit("w620 h420 ReadOnly VScroll HScroll -Wrap", report)
    btn := dbgGui.AddButton("w120 y+10 Default", "Закрыть")
    btn.OnEvent("Click", (*) => dbgGui.Destroy())
    dbgGui.OnEvent("Escape", (*) => dbgGui.Destroy())
    dbgGui.OnEvent("Close", (*) => dbgGui.Destroy())

    GuiResize(guiObj, minMax, w, h) {
        if (minMax = -1) ; свёрнуто в трей — ничего не делаем
            return
        edit.Move(, , w - 30, h - 70)
        btn.Move(, h - 40)
    }
    dbgGui.OnEvent("Size", GuiResize)

    dbgGui.Show("w650 h480")
}

; ------------------------------------------------------------
; Utils() — общие вспомогательные функции
; ------------------------------------------------------------

; превращает "16 800" / "9,5" / "3 066" (в т.ч. с неразрывным пробелом) в число
CleanNum(s) {
    s := StrReplace(s, Chr(0xA0), " ")  ; неразрывный пробел -> обычный, на всякий случай
    s := RegExReplace(s, "\s", "")      ; убираем пробелы-разделители тысяч
    s := StrReplace(s, ",", ".")        ; запятая -> точка (десятичный разделитель)
    return s + 0
}

; форматирует число обратно с пробелами-разделителями тысяч,
; округляя до целого рубля ("4250.7" -> "4 251", копейки не показываем)
FormatNum(n) {
    isNeg := (n < 0)
    n := Round(Abs(n))

    intStr := String(n)
    revStr := ""
    Loop StrLen(intStr)
        revStr .= SubStr(intStr, StrLen(intStr) - A_Index + 1, 1)

    grouped := RegExReplace(revStr, "(\d{3})(?=\d)", "$1 ")
    finalStr := ""
    Loop StrLen(grouped)
        finalStr .= SubStr(grouped, StrLen(grouped) - A_Index + 1, 1)

    return (isNeg ? "-" : "") . finalStr
}
