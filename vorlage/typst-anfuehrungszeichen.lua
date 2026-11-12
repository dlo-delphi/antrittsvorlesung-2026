-- Typst-Ausgabe: Pandoc schreibt das deutsche schließende Anführungszeichen (“)
-- als gerades " in die Typst-Datei. Typst setzt daraus mit lang: de ein
-- öffnendes Zeichen, und aus „…“ wird „…„. Der Filter gibt “ als rohes
-- Typst-Zeichen aus, damit es unverändert im PDF ankommt.
if not FORMAT:match("typst") then
  return {}
end

function Str(el)
  if not el.text:find("“", 1, true) then
    return nil
  end
  local out = pandoc.List()
  local rest = el.text
  while true do
    local i, j = rest:find("“", 1, true)
    if not i then break end
    if i > 1 then out:insert(pandoc.Str(rest:sub(1, i - 1))) end
    out:insert(pandoc.RawInline("typst", "“"))
    rest = rest:sub(j + 1)
  end
  if #rest > 0 then out:insert(pandoc.Str(rest)) end
  return out
end
