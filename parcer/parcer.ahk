#Requires AutoHotkey v2.0

; ============================================
; Преобразование текста КП в таблицу
; Ctrl+Shift+X
; ============================================
^+x::
{
    text := A_Clipboard
    
    if (text = "") {
        MsgBox("Буфер обмена пуст!", "Ошибка", "Icon!")
        return
    }
    
    lines := StrSplit(text, "`n")
    result := []
    
    result.Push("Наименование`tКоличество`tЦена`tСтоимость")
    
    for line in lines {
        line := Trim(line)
        if (line = "")
            continue
        if (RegExMatch(line, "i)^Итого\s+\d"))
            continue
        
        parsed := ParseKPLine(line)
        if (parsed) {
            result.Push(parsed.name "`t" parsed.qty "`t" parsed.price "`t" parsed.total)
        }
    }
    
    ; Собираем строки через цикл
    output := ""
    for i, item in result {
        if (i > 1)
            output .= "`n"
        output .= item
    }
    
    A_Clipboard := output
    
    ToolTip("Текст преобразован в таблицу!")
    SetTimer(() => ToolTip(), -2000)
}

; ============================================
; Преобразование таблицы обратно в текст КП
; Ctrl+Shift+C
; ============================================
^+c::
{
    text := A_Clipboard
    
    if (text = "") {
        MsgBox("Буфер обмена пуст!", "Ошибка", "Icon!")
        return
    }
    
    lines := StrSplit(text, "`n")
    result := []
    
    for line in lines {
        line := Trim(line)
        if (line = "")
            continue
        if (RegExMatch(line, "i)^Наименование"))
            continue
        
        parts := StrSplit(line, "`t")
        
        if (parts.Length >= 4) {
            name := Trim(parts[1])
            qty := Trim(parts[2])
            price := Trim(parts[3])
            total := Trim(parts[4])
            
            result.Push(name " - " qty "шт х " price " руб/шт = " total " руб")
        }
    }
    
    ; Собираем строки через цикл
    output := ""
    for i, item in result {
        if (i > 1)
            output .= "`n"
        output .= item
    }
    
    A_Clipboard := output
    
    ToolTip("Таблица преобразована в текст КП!")
    SetTimer(() => ToolTip(), -2000)
}

; ============================================
; Функция парсинга строки КП
; ============================================
ParseKPLine(line) {
    ; Убираем пробелы-разделители тысяч ("3 066" → "3066")
    line := RegExReplace(line, "(\d)\s+(\d)", "$1$2")
    
    ; Ищем название (всё до первого "-" за которым следует число)
    if (!RegExMatch(line, "^(.+?)\s*-\s*(\d)", &match))
        return 0
    
    name := Trim(match[1])
    rest := SubStr(line, match.Pos(2))
    
    ; Ищем количество
    qty := 1
    if (RegExMatch(rest, "(\d+(?:[.,]\d+)?)\s*шт", &matchQty)) {
        qty := Number(StrReplace(matchQty[1], ",", "."))
    }
    
    ; Ищем цену за штуку
    price := 0
    if (RegExMatch(rest, "(\d+(?:[.,]\d+)?)\s*(?:р|руб)/шт", &matchPrice)) {
        price := Number(StrReplace(matchPrice[1], ",", "."))
    } else if (RegExMatch(rest, "х\s*(\d+(?:[.,]\d+)?)\s*(?:р|руб)", &matchPrice)) {
        price := Number(StrReplace(matchPrice[1], ",", "."))
    } else if (RegExMatch(rest, "-\s*(\d+(?:[.,]\d+)?)\s*(?:р|руб)", &matchPrice)) {
        price := Number(StrReplace(matchPrice[1], ",", "."))
        qty := 1
    }
    
    ; Ищем итоговую стоимость
    total := 0
    if (RegExMatch(rest, "=\s*(\d+(?:[.,]\d+)?)\s*(?:р|руб)", &matchTotal)) {
        total := Number(StrReplace(matchTotal[1], ",", "."))
    } else if (price > 0 && qty > 0) {
        total := Round(price * qty, 2)
    }
    
    if (price = 0 && total = 0)
        return 0
    
    ; Форматируем числа с запятой для русского формата
    qtyStr := FormatNum(qty)
    priceStr := FormatNum(price)
    totalStr := FormatNum(total)
    
    return {name: name, qty: qtyStr, price: priceStr, total: totalStr}
}

; Вспомогательная функция: число → строка с запятой вместо точки
FormatNum(n) {
    if (n = Floor(n)) {
        s := Format("{:d}", Integer(n))
    } else {
        s := Format("{:g}", n)
    }
    return StrReplace(s, ".", ",")
}