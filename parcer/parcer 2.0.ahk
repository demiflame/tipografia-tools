#Requires AutoHotkey v2.0
#SingleInstance Force

; ==========================================
; ЗАЩИТА ОТ ОШИБОК КОДИРОВКИ
; ==========================================
global RUB := "руб"
global SHT := "шт"
global R_SHORT := "р"
global X_CYR := "х"

; ==========================================
; ГОРЯЧИЕ КЛАВИШИ
; ==========================================

^+x:: ; Ctrl + Shift + X
{
    text := A_Clipboard
    if !text
        return
    A_Clipboard := KPToTable(text)
    ToolTip("КП преобразовано в таблицу")
    SetTimer(() => ToolTip(), -1500)
}

^+c:: ; Ctrl + Shift + C
{
    text := A_Clipboard
    if !text
        return
    A_Clipboard := TableToKP(text)
    ToolTip("Таблица преобразована в КП")
    SetTimer(() => ToolTip(), -1500)
}

^+n:: ; Ctrl + Shift + N
{
    text := A_Clipboard
    if !text
        return
    A_Clipboard := ShortList(text)
    ToolTip("Сформирован краткий список")
    SetTimer(() => ToolTip(), -1500)
}

^+d:: ; Ctrl + Shift + D
{
    text := A_Clipboard
    if !text
        return
    Debug(text)
}

; ==========================================
; УТИЛИТЫ (Utils)
; ==========================================

ParseNum(str) {
    str := Trim(str)
    str := StrReplace(str, " ")
    str := StrReplace(str, ",")
    if RegExMatch(str, "^\d+(\.\d+)?$") {
        return Number(str)
    }
    return 0
}

FormatNum(n) {
    n := Round(n)
    str := n 
    formatted := ""
    len := StrLen(str)
    Loop Parse str {
        if (A_Index > 1 && Mod(len - A_Index + 1, 3) == 0)
            formatted .= " "
        formatted .= A_LoopField
    }
    return formatted
}

; ==========================================
; NORMALIZE()
; ==========================================

Normalize(text) {
    text := StrReplace(text, "—", "-")
    text := StrReplace(text, "–", "-")
    text := StrReplace(text, "−", "-")
    
    text := RegExReplace(text, RUB "\s*/\s*" SHT, RUB "/" SHT)
    text := RegExReplace(text, R_SHORT "\s*/\s*" SHT, RUB "/" SHT)
    
    text := RegExReplace(text, "(?<!\d) {2,}(?!\d)", " ")
    return Trim(text)
}

; ==========================================
; TOKENIZER()
; ==========================================

Tokenizer(line) {
    ; Убран флаг "O)", так как он мешал чтению .Pos
    pattern := "\s*-\s*((?:[\d\s,]+\s*" . SHT . "?\s*[" . X_CYR . "x]\s*)?[\d\s,]+\s*" . RUB . "(?:/" . SHT . ")?)\s*$"
    
    ; Убраны внешние скобки вокруг RegExMatch во избежание бага компилятора
    if RegExMatch(line, pattern, &match) {
        name := Trim(SubStr(line, 1, match.Pos - 1))
        calc := Trim(match.Value(1))
        return {name: name, calc: calc}
    }
    
    return {name: line, calc: ""}
}

; ==========================================
; PARSER (Отдельные функции)
; ==========================================

ParseName(nameStr) {
    name := Trim(nameStr)
    name := RegExReplace(name, "\s*-\s*$", "")
    return name
}

ParseQty(calcStr) {
    pattern1 := "(?:^|\s" . X_CYR . "\s*)\s*(\d[\d\s,]*)\s*" . SHT
    pattern2 := "(\d[\d\s,]*)\s*" . SHT . "\s*(?:\s" . X_CYR . "\s|$)"
    
    if RegExMatch(calcStr, pattern1, &m)
        return ParseNum(m.Value(1))
    if RegExMatch(calcStr, pattern2, &m)
        return ParseNum(m.Value(1))
        
    return 1
}

ParsePrice(calcStr) {
    cleanCalc := calcStr
    cleanCalc := RegExReplace(cleanCalc, "\d[\d\s,]*\s*" . SHT . "\s*[" . X_CYR . "x]\s*", "")
    cleanCalc := RegExReplace(cleanCalc, "[" . X_CYR . "x]\s*\d[\d\s,]*\s*" . SHT, "")
    cleanCalc := RegExReplace(cleanCalc, RUB "/" SHT, "")
    cleanCalc := RegExReplace(cleanCalc, RUB, "")
    
    return ParseNum(cleanCalc)
}

ParseTotal(totalStr, qty, price) {
    if (totalStr > 0)
        return totalStr
    return qty * price
}

ParseKPLine(line1, line2 := "") {
    if (line2 != "") {
        line := Trim(line1) . " - " . Trim(line2)
    } else {
        line := line1
    }
    
    line := Normalize(line)
    line := RegExReplace(line, "\s*=\s*[\d\s,]+\s*" . RUB . "\s*$", "")
    
    tokens := Tokenizer(line)
    
    if (tokens.calc == "") {
        return {ok: false, reason: "Не найден блок расчета"}
    }
    
    name := ParseName(tokens.name)
    qty := ParseQty(tokens.calc)
    price := ParsePrice(tokens.calc)
    total := ParseTotal(0, qty, price)
    
    if (price == 0) {
        return {ok: false, reason: "Не удалось определить цену"}
    }
    
    return {ok: true, name: name, qty: qty, price: price, total: total}
}

; ==========================================
; ФОРМАТТЕРЫ И РЕЖИМЫ
; ==========================================

KPToTable(text) {
    lines := StrSplit(text, "`n", "`r")
    result := "Наименование`tКоличество`tЦена`tСтоимость`n"
    priceCheck := RUB "|" R_SHORT "/" SHT
    
    i := 1
    while (i <= lines.Length) {
        line := Trim(lines[i])
        if (line == "") {
            i++
            continue
        }
        
        ; Вынесена сложная проверка в отдельную переменную
        isPriceNext := (i < lines.Length && !RegExMatch(line, priceCheck) && RegExMatch(Trim(lines[i+1]), priceCheck))
        
        if isPriceNext {
            parsed := ParseKPLine(line, lines[i+1])
            i += 2
        } else {
            parsed := ParseKPLine(line)
            i++
        }
        
        if (parsed.ok) {
            result .= parsed.name "`t" parsed.qty " шт`t" FormatNum(parsed.price) " руб`t" FormatNum(parsed.total) " руб`n"
        }
    }
    return Trim(result)
}

TableToKP(text) {
    lines := StrSplit(text, "`n", "`r")
    result := ""
    grandTotal := 0
    
    for i, line in lines {
        line := Trim(line)
        if (line == "" || RegExMatch(line, "^Наименование"))
            continue
            
        cols := StrSplit(line, "`t")
        name := Trim(cols[1])
        
        if (cols.Length == 3) {
            qty := ParseNum(cols[2])
            price := ParseNum(cols[3])
            total := qty * price
        } else if (cols.Length >= 4) {
            qty := ParseNum(cols[2])
            price := ParseNum(cols[3])
            total := ParseNum(cols[4])
        } else {
            continue
        }
        
        grandTotal += total
        result .= name " - " qty "шт х " FormatNum(price) " руб/шт = " FormatNum(total) " руб`n"
    }
    
    result .= "`nИтого к оплате - " FormatNum(grandTotal) " руб`n`nОплата наличными / счет на Юр. лицо"
    return Trim(result)
}

ShortList(text) {
    lines := StrSplit(text, "`n", "`r")
    result := ""
    isTable := RegExMatch(text, "Наименование.*Количество")
    
    if (isTable) {
        for i, line in lines {
            if (RegExMatch(line, "^Наименование") || Trim(line) == "")
                continue
            cols := StrSplit(line, "`t")
            if (cols.Length >= 2) {
                result .= Trim(cols[1]) " - " Trim(cols[2]) "`n"
            }
        }
    } else {
        priceCheck := RUB "|" R_SHORT "/" SHT
        i := 1
        while (i <= lines.Length) {
            line := Trim(lines[i])
            if (line == "") {
                i++
                continue
            }
            
            isPriceNext := (i < lines.Length && !RegExMatch(line, priceCheck) && RegExMatch(Trim(lines[i+1]), priceCheck))
            
            if isPriceNext {
                parsed := ParseKPLine(line, lines[i+1])
                i += 2
            } else {
                parsed := ParseKPLine(line)
                i++
            }
            
            if (parsed.ok) {
                result .= parsed.name " - " parsed.qty " шт`n"
            }
        }
    }
    return Trim(result)
}

Debug(text) {
    lines := StrSplit(text, "`n", "`r")
    totalLines := 0
    successLines := 0
    errors := []
    priceCheck := RUB "|" R_SHORT "/" SHT
    
    i := 1
    while (i <= lines.Length) {
        line := Trim(lines[i])
        if (line == "") {
            i++
            continue
        }
        totalLines++
        
        isPriceNext := (i < lines.Length && !RegExMatch(line, priceCheck) && RegExMatch(Trim(lines[i+1]), priceCheck))
        
        if isPriceNext {
            parsed := ParseKPLine(line, lines[i+1])
            i += 2
        } else {
            parsed := ParseKPLine(line)
            i++
        }
        
        if (parsed.ok) {
            successLines++
        } else {
            errors.Push("Строка: " line "`nПричина: " parsed.reason)
        }
    }
    
    msg := "=== ДИАГНОСТИКА ПАРСЕРА ===`n`n"
    msg .= "Всего строк: " totalLines "`n"
    msg .= "Распознано: " successLines "`n"
    msg .= "Не распознано: " (totalLines - successLines) "`n"
    
    if (errors.Length > 0) {
        msg .= "`n--- ОШИБКИ ---`n"
        for e in errors {
            msg .= e "`n`n"
        }
    }
    
    Gui("Destroy")
    gui := Gui("+Resize", "Отчет диагностики")
    gui.AddEdit("r20 w500 vMsg", msg).ReadOnly := true
    gui.Show()
}