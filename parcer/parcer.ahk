#Requires AutoHotkey v2.0

global DebugTotal := 0
global DebugParsed := 0
global DebugErrors := ""

; ============================================
; CTRL+SHIFT+X
; КП -> Таблица
; ============================================
^+x::
{
    global DebugTotal, DebugParsed, DebugErrors

    DebugTotal := 0
    DebugParsed := 0
    DebugErrors := ""

    text := A_Clipboard

    if (Trim(text) = "")
    {
        MsgBox("Буфер обмена пуст!")
        return
    }

    result := []

    lines := StrSplit(text, "`n", "`r")

    pendingName := ""

    for line in lines
    {
        line := Trim(line)

        if (line = "")
            continue

        if RegExMatch(line, "i)^итого")
            continue

        DebugTotal++

        ; двухстрочная позиция
        if !RegExMatch(line, "i)(шт|руб|р/шт|/шт)")
        {
            pendingName := line
            continue
        }

        if (pendingName != "" && !InStr(line, "-"))
        {
            line := pendingName " - " line
            pendingName := ""
        }

        parsed := ParseKPLine(line)

        if IsObject(parsed)
        {
            result.Push(
                parsed.name "`t"
                parsed.qty "`t"
                parsed.price "`t"
                parsed.total
            )

            DebugParsed++
        }
        else
        {
            DebugErrors .= line "`r`n"
        }
    }

    output := ""

    for i, row in result
    {
        if (i > 1)
            output .= "`r`n"

        output .= row
    }

    A_Clipboard := output

    ToolTip("КП -> Таблица")
    SetTimer(() => ToolTip(), -1500)
}

; ============================================
; CTRL+SHIFT+C
; Таблица -> КП
; ============================================
^+c::
{
    text := A_Clipboard

    if (Trim(text) = "")
    {
        MsgBox("Буфер обмена пуст!")
        return
    }

    lines := StrSplit(text, "`n", "`r")

    result := []

    grandTotal := 0

    for line in lines
    {
        line := Trim(line)

        if (line = "")
            continue

        if RegExMatch(line, "i)^наименование")
            continue

        parts := StrSplit(line, "`t")

        if (parts.Length < 3)
            continue

        name := Trim(parts[1])
        qty := Trim(parts[2])
        price := Trim(parts[3])

        if (parts.Length >= 4 && Trim(parts[4]) != "")
        {
            total := Trim(parts[4])
        }
        else
        {
            q := ToNumber(qty)
            p := ToNumber(price)

            total := FormatNum(Round(q * p, 2))
        }

        grandTotal += ToNumber(total)

        result.Push(
            name
            . " - "
            . qty
            . "шт х "
            . price
            . " руб/шт = "
            . total
            . " руб"
        )
    }

    output := ""

    for i, row in result
    {
        if (i > 1)
            output .= "`r`n"

        output .= row
    }

    output .= "`r`n`r`n"
    output .= "Итого к оплате - "
    output .= FormatNum(grandTotal)
    output .= " руб"
    output .= "`r`n`r`n"
    output .= "Оплата наличными / счет на Юр. лицо"

    A_Clipboard := output

    ToolTip("Таблица -> КП")
    SetTimer(() => ToolTip(), -1500)
}

; ============================================
; CTRL+SHIFT+D
; Отладка
; ============================================
^+d::
{
    global DebugTotal, DebugParsed, DebugErrors

    msg :=
    (
    "Всего строк: " DebugTotal "`n"
    "Распознано: " DebugParsed "`n`n"
    "Не распознано:`n`n"
    DebugErrors
    )

    MsgBox(msg)
}

; ============================================
; ПАРСЕР
; ============================================
ParseKPLine(line)
{
    ; убрать пробелы в тысячах
    line := RegExReplace(line, "(\d)\s+(\d)", "$1$2")
line := StrReplace(line, "—", "-")
line := StrReplace(line, "–", "-")
    ; --------------------------------
    ; ручка 85 руб/шт х 50 шт = 4250 руб
    ; --------------------------------

    if RegExMatch(
        line,
        "^(.*?)\s+(\d+(?:[.,]\d+)?)\s*(?:р(?:уб)?\s*/?\s*шт|р/шт)\s*[хx]\s*(\d+)\s*шт(?:\s*=\s*(\d+(?:[.,]\d+)?))?",
        &m)
    {
        return MakeResult(
            Trim(m[1]),
            m[3],
            m[2],
            m[4]
        )
    }

    ; --------------------------------
    ; название - 10шт х 420р/шт
    ; --------------------------------

    if RegExMatch(
        line,
        "^(.*?)\s*-\s*(\d+)\s*шт\s*[хx]\s*(\d+(?:[.,]\d+)?)\s*(?:р(?:уб)?\s*/?\s*шт|р/шт)(?:\s*=\s*(\d+(?:[.,]\d+)?))?",
        &m)
    {
        return MakeResult(
            Trim(m[1]),
            m[2],
            m[3],
            m[4]
        )
    }

    ; --------------------------------
    ; название - 420р/шт х 10шт
    ; --------------------------------

    if RegExMatch(
        line,
        "^(.*?)\s*-\s*(\d+(?:[.,]\d+)?)\s*(?:р(?:уб)?\s*/?\s*шт|р/шт)\s*[хx]\s*(\d+)\s*шт(?:\s*=\s*(\d+(?:[.,]\d+)?))?",
        &m)
    {
        return MakeResult(
            Trim(m[1]),
            m[3],
            m[2],
            m[4]
        )
    }

    ; --------------------------------
    ; ролл ап - 6300 руб/шт
    ; --------------------------------

    if RegExMatch(
        line,
        "^(.*?)\s*-\s*(\d+(?:[.,]\d+)?)\s*(?:р(?:уб)?\s*/?\s*шт|р/шт)$",
        &m)
    {
        return MakeResult(
            Trim(m[1]),
            1,
            m[2],
            m[2]
        )
    }

    ; --------------------------------
    ; А4 - 50руб
    ; --------------------------------

    if RegExMatch(
        line,
        "^(.*?)\s*-\s*(\d+(?:[.,]\d+)?)\s*(?:р|руб)$",
        &m)
    {
        return MakeResult(
            Trim(m[1]),
            1,
            m[2],
            m[2]
        )
    }

    return 0
}

; ============================================
; СОЗДАНИЕ ОБЪЕКТА
; ============================================
MakeResult(name, qty, price, total)
{
    qtyNum := ToNumber(qty)
    priceNum := ToNumber(price)

    if (total = "")
        totalNum := Round(qtyNum * priceNum, 2)
    else
        totalNum := ToNumber(total)

    return {
        name: name,
        qty: FormatNum(qtyNum),
        price: FormatNum(priceNum),
        total: FormatNum(totalNum)
    }
}

; ============================================
; ЧИСЛО ИЗ СТРОКИ
; ============================================
ToNumber(val)
{
    val := Trim(val)
    val := StrReplace(val, " ", "")
    val := StrReplace(val, ",", ".")

    return val + 0
}

; ============================================
; ФОРМАТИРОВАНИЕ
; ============================================
FormatNum(n)
{
    if (n = Floor(n))
        return Integer(n)

    return StrReplace(Format("{:g}", n), ".", ",")
}