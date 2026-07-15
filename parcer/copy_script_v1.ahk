#Requires AutoHotkey v2.0
#SingleInstance Force

; ============================================================
;  Parser v2.0 — универсальный парсер коммерческих предложений
;  Работает через буфер обмена (только чтение/запись A_Clipboard,
;  без автокопирования и автовставки — вставляет пользователь сам)
;
;  Ctrl+Shift+X — КП -> Таблица
;  Ctrl+Shift+C — Таблица -> КП
;  Ctrl+Shift+D — Диагностика (MsgBox)
;  Ctrl+Shift+N — Краткий список "Наименование - Количество"
; ============================================================

TrayTip("Parser v2.0", "Скрипт запущен и готов к работе.")

; ------------------------------------------------------------
; Hotkeys
; ------------------------------------------------------------

^+x:: KPToTable()
^+c:: TableToKP()
^+d:: DebugMode()
^+n:: ShortList()

; ------------------------------------------------------------
; Normalize() — нормализация исходного текста
; ------------------------------------------------------------

Normalize(text) {
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
; ------------------------------------------------------------

Tokenizer(line) {
    line := Trim(line)
    if (line = "")
        return {name: "", calc: ""}

    ; --- Проход 1: ищем последний разделитель " - " ---
    lastDashPos := 0
    searchPos := 1
    while (foundPos := InStr(line, " - ", , searchPos)) {
        lastDashPos := foundPos
        searchPos := foundPos + 1
    }
    if (lastDashPos) {
        namePart := SubStr(line, 1, lastDashPos - 1)
        calcPart := SubStr(line, lastDashPos + 3)
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

; ------------------------------------------------------------
; Parser — отдельные функции разбора
; ------------------------------------------------------------

ParseName(nameStr) {
    return Trim(nameStr, " -,`t")
}

ParseQty(calc) {
    if RegExMatch(calc, "i)(" . NUM . ")\s*шт", &m)
        return CleanNum(m[1])
    return 1
}

ParsePrice(calc) {
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

    ; 4) запасной вариант — просто число
    if RegExMatch(calc, "(" . NUM . ")", &m4)
        return CleanNum(m4[1])

    return 0
}

ParseTotal(calc, qty, price) {
    if RegExMatch(calc, "=\s*(" . NUM . ")", &m)
        return CleanNum(m[1])
    return qty * price
}

; собирает одну позицию из строки (название + расчёт уже объединены через " - ")
ParseKPLine(line) {
    line := Trim(line)
    if (line = "")
        return {ok: false, error: "Пустая строка", raw: line}

    tok := Tokenizer(line)
    if (tok.calc = "")
        return {ok: false, error: "Не найдена расчётная часть (цена/шт)", raw: line}

    name := ParseName(tok.name)
    if (name = "")
        return {ok: false, error: "Не найдено название", raw: line}

    qty := ParseQty(tok.calc)
    price := ParsePrice(tok.calc)
    if (price = 0)
        return {ok: false, error: "Не удалось распознать цену", raw: line}

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
        } else if (tok.name != "" && tok.calc = "" && i < n) {
            ; возможно двухстрочный формат: название + расчёт на след. строке
            tok2 := Tokenizer(lines[i + 1])
            if (tok2.name = "" && tok2.calc != "") {
                combined := tok.name . " - " . tok2.calc
                items.Push(ParseKPLine(combined))
                i += 2
            } else {
                items.Push({ok: false, error: "Не найдена расчётная часть (цена/шт)", raw: line})
                i += 1
            }
        } else {
            items.Push({ok: false, error: "Не удалось разобрать строку", raw: line})
            i += 1
        }
    }
    return items
}

; ------------------------------------------------------------
; Сборка позиций из таблицы (Tab-разделённые колонки)
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
        if (cols.Length < 3)
            continue

        name := Trim(cols[1])
        qty := CleanNum(cols[2])
        price := CleanNum(cols[3])
        if (cols.Length >= 4 && Trim(cols[4]) != "")
            total := CleanNum(cols[4])
        else
            total := qty * price

        items.Push({ok: true, name: name, qty: qty, price: price, total: total})
    }
    return items
}

; ------------------------------------------------------------
; КП -> Таблица (Ctrl+Shift+X)
; ------------------------------------------------------------

KPToTable() {
    text := A_Clipboard
    if (Trim(text) = "") {
        MsgBox("Буфер обмена пуст.", "Parser v2.0", "Icon!")
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
        MsgBox("Не удалось распознать ни одной строки.`nИспользуйте Ctrl+Shift+D для диагностики.", "Parser v2.0", "Icon!")
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
    text := A_Clipboard
    if (Trim(text) = "") {
        MsgBox("Буфер обмена пуст.", "Parser v2.0", "Icon!")
        return
    }

    items := BuildItemsFromTable(text)
    if (items.Length = 0) {
        MsgBox("Не удалось распознать таблицу.`nОжидается формат: Наименование [Tab] Количество [Tab] Цена [Tab] Стоимость (Стоимость необязательна).", "Parser v2.0", "Icon!")
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
    ToolTip("Готово: КП сформировано (" . items.Length . " позиций). Результат в буфере обмена.")
    SetTimer(() => ToolTip(), -2500)
}

; ------------------------------------------------------------
; Краткий список (Ctrl+Shift+N)
; ------------------------------------------------------------

ShortList() {
    text := A_Clipboard
    if (Trim(text) = "") {
        MsgBox("Буфер обмена пуст.", "Parser v2.0", "Icon!")
        return
    }

    items := []
    if (InStr(text, "`t") && RegExMatch(text, "i)Наименование")) {
        items := BuildItemsFromTable(text)
    } else {
        for _, it in BuildItemsFromKP(text) {
            if (it.ok)
                items.Push(it)
        }
    }

    if (items.Length = 0) {
        MsgBox("Не удалось распознать ни одной позиции.", "Parser v2.0", "Icon!")
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
; Диагностика (Ctrl+Shift+D)
; ------------------------------------------------------------

DebugMode() {
    text := A_Clipboard
    if (Trim(text) = "") {
        MsgBox("Буфер обмена пуст.", "Parser v2.0", "Icon!")
        return
    }

    items := BuildItemsFromKP(text)
    okList := []
    failList := []
    for _, it in items {
        if (it.ok)
            okList.Push(it)
        else
            failList.Push(it)
    }

    msg := "=== Диагностика Parser v2.0 ===`n`n"
    msg .= "Всего строк/позиций: " . items.Length . "`n"
    msg .= "Распознано успешно: " . okList.Length . "`n"
    msg .= "Не распознано: " . failList.Length . "`n"

    if (failList.Length > 0) {
        msg .= "`n--- Проблемные строки ---`n"
        for _, f in failList
            msg .= "`n[" . f.raw . "]`nПричина: " . f.error . "`n"
    }

    MsgBox(msg, "Parser v2.0 — Диагностика", "Iconi")
}

; ------------------------------------------------------------
; Utils() — общие вспомогательные функции и константы
; ------------------------------------------------------------

; число с опциональной группировкой пробелами (тысячи) и десятичной частью через запятую/точку
NUM := "\d{1,3}(?:[ ]\d{3})*(?:[.,]\d+)?"

; строка "похожа на расчётную часть", если содержит признаки цены/количества
LooksLikeCalc(s) {
    return RegExMatch(s, "i)(руб|р\s*/\s*шт|\bшт\b)") ? true : false
}

; превращает "16 800" / "9,5" / "3 066" в число
CleanNum(s) {
    s := RegExReplace(s, "\s", "")   ; убираем пробелы-разделители тысяч
    s := StrReplace(s, ",", ".")     ; запятая -> точка (десятичный разделитель)
    return s + 0
}

; форматирует число обратно с пробелами-разделителями тысяч ("4250" -> "4 250")
FormatNum(n) {
    isNeg := (n < 0)
    n := Abs(n)
    intPart := Floor(n)
    decPart := Round((n - intPart) * 100)

    intStr := String(intPart)
    revStr := ""
    Loop StrLen(intStr)
        revStr .= SubStr(intStr, StrLen(intStr) - A_Index + 1, 1)

    grouped := RegExReplace(revStr, "(\d{3})(?=\d)", "$1 ")
    finalStr := ""
    Loop StrLen(grouped)
        finalStr .= SubStr(grouped, StrLen(grouped) - A_Index + 1, 1)

    result := (isNeg ? "-" : "") . finalStr
    if (decPart > 0)
        result .= "," . Format("{:02}", decPart)

    return result
}
