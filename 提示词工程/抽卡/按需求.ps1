#Requires -Version 5.1
<#
  按需求出词 - 根据一句要求或关键词，从词库匹配提示词，输出正向 + 反向
  支持整句中文/英文：自动扫描句子中出现的词条，并按空格/逗号拆分关键词
  数据来源: 上级目录 提示词对照.xml

  用法示例:
    .\按需求.ps1 "一个银发双马尾的猫耳女孩，穿着泳装，在泳池边害羞地笑着"
    .\按需求.ps1 "银发 双马尾 猫耳 泳装 害羞"
    .\按需求.ps1 "silver hair cat ears swimsuit 校园"
    .\按需求.ps1 "触手 束缚" -NegCount 20 -Copy

  特殊规则：女性向 / 乙游 / BL 游戏·全男系列的角色（名单见 共享常量.json 的 goreOnly 与 抽卡\女性向男同角色.txt）
            命中后只输出血腥暴力提示词（含杀害 / 秀色食人），不做情色拆解。
            例：.\按需求.ps1 "加州清光" 或 "刀剑乱舞"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)][string]$Query,
    [int]$NegCount = 0,          # 大于 0 时用随机反向词；等于 0 用内置核心反向词
    [switch]$Copy,
    [string]$Out
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$root    = Split-Path -Parent $MyInvocation.MyCommand.Path

# ---- 共享常量（单一来源：共享常量.json；网页版由 _build\build-constants.ps1 生成 共享常量.js） ----
$scPath = Join-Path $root '共享常量.json'
if (-not (Test-Path -LiteralPath $scPath)) { $scPath = Join-Path (Split-Path -Parent $root) '共享常量.json' }
if (-not (Test-Path -LiteralPath $scPath)) { throw '找不到 共享常量.json，请放在本脚本目录或其上级目录' }
$SC = Get-Content -LiteralPath $scPath -Raw -Encoding UTF8 | ConvertFrom-Json

# 黑名单词条（共享常量.json 的 wipeTerms）：名牌 / 写名字的框之类，任何模式都不进正向（负面词类别不受影响）
$wipeEn = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
if ($SC.wipeTerms) { foreach ($w in @($SC.wipeTerms.terms)) { if ($w) { [void]$wipeEn.Add(([string]$w).Trim()) } } }

$names   = @('提示词总库.xml','提示词对照.xml')
$xmlPath = $null
foreach ($n in $names) { $p = Join-Path $root $n; if (Test-Path -LiteralPath $p) { $xmlPath = $p; break } }
if (-not $xmlPath) { foreach ($n in $names) { $p = Join-Path (Split-Path -Parent $root) $n; if (Test-Path -LiteralPath $p) { $xmlPath = $p; break } } }
if (-not $xmlPath) { throw "找不到词库（提示词总库.xml 或 提示词对照.xml），请放在本脚本目录或其上级目录" }

$L_cat = New-Object System.Collections.Generic.List[string]
$L_id  = New-Object System.Collections.Generic.List[string]
$L_en  = New-Object System.Collections.Generic.List[string]
$L_zh  = New-Object System.Collections.Generic.List[string]
$L_enL = New-Object System.Collections.Generic.List[string]
$L_zhL = New-Object System.Collections.Generic.List[string]
$negTerms = New-Object System.Collections.Generic.List[object]
$goreTerms = New-Object System.Collections.Generic.List[object]   # gore-combos（血腥暴力组合）单独收集，不参与逐词匹配
$catIdx = @{}                                                     # 类别 id -> 该类别词条在 L_* 里的下标（联动词取词用）

# 流式读取（不建 DOM），跳过 erotic-combos：整段组合对逐词匹配无意义，且在 20MB+ 时会明显拖慢
$rs = New-Object System.Xml.XmlReaderSettings
$rs.IgnoreComments = $true
$rs.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
$rd = [System.Xml.XmlReader]::Create($xmlPath, $rs)
try {
    $curZh = ''; $curId = ''; $skip = $false
    while ($rd.Read()) {
        if ($rd.NodeType -eq [System.Xml.XmlNodeType]::Element) {
            if ($rd.Name -eq 'category') {
                $cid = $rd.GetAttribute('id'); $czh = $rd.GetAttribute('zh')
                if ($cid -eq 'erotic-combos') { $skip = $true; $curId = ''; $inGore = $false }
                elseif ($cid -eq 'gore-combos') { $skip = $true; $curId = ''; $inGore = $true }
                else { $skip = $false; $curId = $cid; $curZh = $czh; $inGore = $false }
            }
            elseif ($rd.Name -eq 'term') {
                if ($inGore) { $goreTerms.Add([pscustomobject]@{ en = $rd.GetAttribute('en'); zh = $rd.GetAttribute('zh') }) }
                elseif (-not $skip -and $curId) {
                    $en = $rd.GetAttribute('en'); $zh = $rd.GetAttribute('zh')
                    if ($curId -eq 'negative') { $negTerms.Add([pscustomobject]@{ en = $en; zh = $zh }) }
                    elseif (-not $wipeEn.Contains(([string]$en).Trim())) {   # 黑名单词条不进正向
                        $ixNow = $L_en.Count
                        $L_cat.Add($curZh); $L_id.Add($curId); $L_en.Add($en); $L_zh.Add($zh)
                        $L_enL.Add($en.ToLowerInvariant()); $L_zhL.Add($zh.ToLowerInvariant())
                        if (-not $catIdx.ContainsKey($curId)) { $catIdx[$curId] = New-Object System.Collections.Generic.List[int] }
                        $catIdx[$curId].Add($ixNow)
                    }
                }
            }
        }
        elseif ($rd.NodeType -eq [System.Xml.XmlNodeType]::EndElement -and $rd.Name -eq 'category') { $skip = $false; $curId = ''; $inGore = $false }
    }
}
finally { $rd.Close() }

$CORE = @($SC.core)   # 反向核心词（来源: 共享常量.json）

# ===== 女性向 / 乙游 / BL 游戏·全男系列：按设定只输出血腥暴力提示词（规则来源: 共享常量.json 的 goreOnly） =====
$GORE = $SC.goreOnly
$GoreOn = [bool]($GORE -and $GORE.enabled)
$gorePat = @()
if ($GoreOn) {
    $gorePat += @($GORE.series)
    if ($GORE.useAuto) { foreach ($gk in @($GORE.seriesAuto)) { $gorePat += ($gk + '*') } }
}
$GoreCount = $(if ($GoreOn -and $GORE.count) { [int]$GORE.count } else { 1 })
$GoreWithChar = [bool]($GoreOn -and $GORE.withCharacter)
function Test-GoreSeries([string]$s) {
    if (-not $GoreOn -or -not $s) { return $false }
    $k = (((([string]$s) -replace '\\', '') -replace '\s+', ' ').Trim().ToLowerInvariant())
    foreach ($gp in $gorePat) { if ($k -like $gp) { return $true } }
    return $false
}
# 从 character 词条里取系列名（尾部括号），再判是否属于这类作品
function Get-SeriesOfTag([string]$en) {
    if ($en -match '\(([^()]+)\)\s*$') { return $matches[1] }
    return ''
}
# 整句里是否出现某个 tag（词边界，避免子串误命中）
function Test-WordIn([string]$q, [string]$k) {
    if (-not $q -or -not $k) { return $false }
    $kk = $k.ToLowerInvariant()
    if (-not $q.Contains($kk)) { return $false }
    return [regex]::IsMatch($q, '(^|[^a-z0-9])' + [regex]::Escape($kk) + '([^a-z0-9]|$)')
}
# 系列名单（去掉通配符 *）→ 直接拿整句比对
$goreKeys = New-Object System.Collections.Generic.List[string]
foreach ($gp in $gorePat) { $goreKeys.Add((($gp -replace '\*','').Trim())) }
# 角色名单：抽卡\女性向男同角色.txt（由 _build\build-gore-chars.ps1 生成）
$blChars = @{}
$blCharKeys = New-Object System.Collections.Generic.List[string]
if ($GoreOn) {
    foreach ($bn in @('女性向男同角色.txt')) {
        $bp = Join-Path $root $bn
        if (-not (Test-Path -LiteralPath $bp)) { $bp = Join-Path (Split-Path -Parent $root) $bn }
        if (Test-Path -LiteralPath $bp) {
            foreach ($bl in [System.IO.File]::ReadAllLines($bp, [System.Text.Encoding]::UTF8)) {
                if (-not $bl -or $bl.StartsWith('#')) { continue }
                $pp = $bl -split "`t"
                if ($pp.Count -ge 1 -and $pp[0].Trim()) { $blChars[$pp[0].Trim()] = $true; $blCharKeys.Add($pp[0].Trim()) }
            }
            break
        }
    }
}
# 词库不含 gore-combos 时（例如只找到 提示词对照.xml）单独加载 血腥暴力组合库.xml
function Read-GoreCombos([string]$path) {
    $rs2 = New-Object System.Xml.XmlReaderSettings
    $rs2.IgnoreComments = $true; $rs2.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
    $rd2 = [System.Xml.XmlReader]::Create($path, $rs2)
    try {
        $inG = $false
        while ($rd2.Read()) {
            if ($rd2.NodeType -eq [System.Xml.XmlNodeType]::Element) {
                if ($rd2.Name -eq 'category') { $inG = ($rd2.GetAttribute('id') -eq 'gore-combos') }
                elseif ($inG -and $rd2.Name -eq 'term') { $goreTerms.Add([pscustomobject]@{ en = $rd2.GetAttribute('en'); zh = $rd2.GetAttribute('zh') }) }
            }
        }
    }
    finally { $rd2.Close() }
}

# ===== 换装：服装/制服类词判定（把角色卡里的原服装词剔除后，换上你点名的新服装） =====
$WearIds = New-Object 'System.Collections.Generic.HashSet[string]'
@($SC.wearIds) | ForEach-Object { [void]$WearIds.Add($_) }   # 服装/制服类（来源: 共享常量.json）
$WearEn = New-Object 'System.Collections.Generic.HashSet[string]'
$WearZh = New-Object 'System.Collections.Generic.HashSet[string]'
for ($wi = 0; $wi -lt $L_id.Count; $wi++) {
    if ($WearIds.Contains($L_id[$wi])) { [void]$WearEn.Add($L_enL[$wi]); [void]$WearZh.Add($L_zhL[$wi]) }
}
$WearKw = [string]$SC.wearKw   # 服装关键词（来源: 共享常量.json）
function Test-WearToken { param([string]$t) $l = ([string]$t).ToLower().Trim(); if (-not $l) { return $false }; return ($WearEn.Contains($l) -or $WearZh.Contains($l) -or ($l -match $WearKw)) }
function Remove-WearToken { param([string]$en) $keep = @(([string]$en) -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -and -not (Test-WearToken $_) }); return ($keep -join ', ') }

# 特征词归一化：去转义反斜杠、下划线转空格、压缩空白、转小写（与词库 en 对齐）
function NormFeat { param([string]$s) return (((([string]$s) -replace '\\','') -replace '_',' ') -replace '\s+',' ').Trim().ToLowerInvariant() }

# 反向词
if ($NegCount -gt 0 -and $negTerms) {
    $pool = @(); foreach ($t in $negTerms) { $pool += [pscustomobject]@{ en = $t.en; zh = $t.zh } }
    $rand = New-Object System.Random
    for ($i = $pool.Count - 1; $i -gt 0; $i--) { $j = $rand.Next($i + 1); $tmp = $pool[$i]; $pool[$i] = $pool[$j]; $pool[$j] = $tmp }
    $negList = @($pool | Select-Object -First ([Math]::Min($NegCount, $pool.Count)))
} else {
    $negList = @($CORE | ForEach-Object { [pscustomobject]@{ en = $_; zh = '' } })
}

# 中文俗名 -> 词库标准中文名（丝袜不写种类时一律按“连裤袜”）——来源: 共享常量.json
$SYN = @{}; $synKeys = New-Object System.Collections.Generic.List[string]
foreach ($p in $SC.syn) { $SYN[[string]$p[0]] = [string]$p[1]; $synKeys.Add([string]$p[0]) }   # 源序来自 共享常量.json
$sLower = $Query.ToLower()

# 0) 角色卡：命中即把整条复刻提示词加入结果；其余文字继续逐词拆解并追加
$cardList = @()
foreach ($cp in @((Join-Path $root '角色卡.js'), (Join-Path (Split-Path -Parent $root) '角色卡.js'))) {
    if (Test-Path -LiteralPath $cp) {
        $ct = [System.IO.File]::ReadAllText($cp, [System.Text.Encoding]::UTF8)
        $bi = $ct.IndexOf('['); $bj = $ct.LastIndexOf(']')
        if ($bi -ge 0 -and $bj -gt $bi) {
            $json = $ct.Substring($bi, $bj - $bi + 1)
            # JavaScriptSerializer 比自带 ConvertFrom-Json 快约 2 倍（依赖 .NET Framework 的 System.Web.Extensions）
            try {
                Add-Type -AssemblyName System.Web.Extensions -ErrorAction Stop
                $ser = New-Object System.Web.Script.Serialization.JavaScriptSerializer
                $ser.MaxJsonLength = [int]::MaxValue
                $cardList = @($ser.DeserializeObject($json))
            } catch {
                try { $cardList = @(($json | ConvertFrom-Json)) } catch { $cardList = @() }
            }
        }
        break
    }
}
$cardHit = New-Object System.Collections.Generic.List[object]
$hits = New-Object System.Collections.Generic.List[object]
foreach ($cd in $cardList) {
    $keys = @($cd.name_en, $cd.name_zh, $cd.id) + @($cd.aliases)
    $best = ''
    foreach ($k in $keys) { $kn = ([string]$k).ToLower().Trim(); if ($kn.Length -ge 2 -and $sLower.Contains($kn) -and $kn.Length -gt $best.Length) { $best = $kn } }
    if ($best) { $hits.Add([pscustomobject]@{ card = $cd; best = $best }) }
}
# 命中的别名越长越具体优先；并列时 name_en 括号少者（角色本体）优先
$hits = @($hits | Sort-Object -Property @{ Expression = { $_.best.Length }; Descending = $true },
                                       @{ Expression = { (([string]$_.card.name_en) -split '\(').Count }; Descending = $false })
$kept = New-Object System.Collections.Generic.List[object]
for ($hi = 0; $hi -lt $hits.Count; $hi++) {
    $h = $hits[$hi]
    if (Test-WearToken $h.best) {
        $stronger = $false
        for ($oi = 0; $oi -lt $hits.Count; $oi++) {
            if ($oi -eq $hi) { continue }
            $o = $hits[$oi]
            if ((-not (Test-WearToken $o.best)) -and $o.best.Length -gt $h.best.Length) { $stronger = $true; break }
        }
        if ($stronger) { continue }   # 只是被普通服装词命中，且有更具体的角色 → 当服装处理
    }
    $dup = $false
    foreach ($k in $kept) { if ($k.best -eq $h.best -or $k.best.Contains($h.best)) { $dup = $true; break } }
    if (-not $dup) { $kept.Add($h) }
}
foreach ($h in $kept) { $cardHit.Add($h.card) }

$posList = New-Object System.Collections.Generic.List[object]
$seen = @{}
$groupLines = @()
$featLines = @()
$cardTokens = @{}
$scanStr = $sLower

if ($cardHit.Count -gt 0) {
    foreach ($cd in $cardHit) {
        $groupLines += ("  [角色卡] {0} {1}" -f $cd.name_zh, $cd.name_en)
        foreach ($tk in ([string]$cd.prompt_en -split ',')) { $tt = $tk.Trim().ToLower(); if ($tt) { $cardTokens[$tt] = $true } }
        $rmKeys = @(@($cd.name_en, $cd.name_zh, $cd.id) + @($cd.aliases)) | Sort-Object { ([string]$_).Length } -Descending
        foreach ($k in $rmKeys) {
            $kn = ([string]$k).ToLower().Trim(); if ($kn.Length -ge 2) { $scanStr = $scanStr.Replace($kn, ' ') }
        }
    }
}


# ===== 规则检查：女性向 / 乙游 / BL 游戏·全男系列的角色 / 作品（用完整查询匹配，不受角色卡名剔除影响）=====
$goreCharEn = ''; $goreCharZh = ''; $goreSeriesEn = ''; $goreSeriesZh = ''
$goreOnly = $false
$goreWarn = ''
if ($GoreOn) {
    for ($gi0 = 0; $gi0 -lt $L_id.Count; $gi0++) {
        $lid0 = $L_id[$gi0]
        if ($lid0 -eq 'character' -and $goreCharEn) { continue }
        if ($lid0 -eq 'copyright' -and $goreSeriesEn) { continue }
        if ($lid0 -ne 'character' -and $lid0 -ne 'copyright') { continue }
        $hit0 = $false
        $zl0 = $L_zhL[$gi0]
        if ($zl0.Length -ge 2 -and $sLower.Contains($zl0)) {
            if ($zl0 -match '[\u4e00-\u9fff\u3040-\u30ff\uac00-\ud7af]') { $hit0 = $true }
            elseif ([regex]::IsMatch($sLower, '(^|[^a-z0-9])' + [regex]::Escape($zl0) + '([^a-z0-9]|$)')) { $hit0 = $true }
        }
        if (-not $hit0) {
            $el0 = $L_enL[$gi0]
            if ($el0.Length -ge 3 -and $sLower.Contains($el0)) {
                if ([regex]::IsMatch($sLower, '(^|[^a-z])' + [regex]::Escape($el0) + '([^a-z]|$)')) { $hit0 = $true }
            }
        }
        if (-not $hit0) { continue }
        $en0 = $L_en[$gi0]
        if ($lid0 -eq 'character') {
            # 角色关联的系列名在词库里查不到 → 查名单文件
            if ($blChars.ContainsKey($en0) -or (Test-GoreSeries (Get-SeriesOfTag $en0))) { $goreCharEn = $en0; $goreCharZh = $L_zh[$gi0] }
        }
        else {
            if (Test-GoreSeries $en0) { $goreSeriesEn = $en0; $goreSeriesZh = $L_zh[$gi0] }
        }
    }
    # 名单直配整句：填上词库没覆盖到的角色 / 系列（只给英文 tag，中文名从词库命中时才有）
    if (-not $goreCharEn) {
        foreach ($bk in $blCharKeys) { if (Test-WordIn $sLower $bk) { $goreCharEn = $bk; break } }
    }
    if (-not $goreSeriesEn) {
        foreach ($bk in $goreKeys) { if (Test-WordIn $sLower $bk) { $goreSeriesEn = $bk; break } }
    }
    # 命中即锁定：不输出情色内容，只出血腥暴力
    if ($goreCharEn -or $goreSeriesEn) { $goreOnly = $true; $cardHit.Clear(); $groupLines = @() }
}

# 1) 整句扫描（命中角色卡时，扫描的是去掉角色名后的剩余文字，作为“附加拆解”追加在卡片之后）
if (-not $goreOnly) {
for ($i = 0; $i -lt $L_zhL.Count; $i++) {
    $m = $false
    $zl = $L_zhL[$i]
    if ($zl.Length -ge 2 -and $scanStr.Contains($zl)) {
        if ($zl -match '[\u4e00-\u9fff\u3040-\u30ff\uac00-\ud7af]') { $m = $true }
        elseif ([regex]::IsMatch($scanStr, '(^|[^a-z0-9])' + [regex]::Escape($zl) + '([^a-z0-9]|$)')) { $m = $true }
    }
    if (-not $m) {
        $el = $L_enL[$i]
        if ($el.Length -ge 3 -and $scanStr.Contains($el)) {
            $pat = '(^|[^a-z])' + [regex]::Escape($el) + '([^a-z]|$)'
            if ([regex]::IsMatch($scanStr, $pat)) { $m = $true }
        }
    }
    if ($m) {
        $en = $L_en[$i]
        if (-not $seen.ContainsKey($en) -and -not $cardTokens.ContainsKey($en.ToLower())) {
            $seen[$en] = $true; $posList.Add([pscustomobject]@{ cat = $L_cat[$i]; id = $L_id[$i]; en = $en; zh = $L_zh[$i] })
        }
    }
}
}

# ===== 命中女性向 / 乙游 / BL 游戏·全男系列的角色 → 按设定只输出血腥暴力提示词 =====
if ($goreOnly) {
    $posList.Clear()
    $seen = @{}
    if ($goreTerms.Count -eq 0) {
        foreach ($gfn in @('血腥暴力组合库.xml')) {
            $gp2 = Join-Path $root $gfn
            if (-not (Test-Path -LiteralPath $gp2)) { $gp2 = Join-Path (Split-Path -Parent $root) $gfn }
            if (Test-Path -LiteralPath $gp2) { Read-GoreCombos $gp2; break }
        }
    }
    if ($GoreWithChar) {
        $gtagEn = $(if ($goreCharEn) { $goreCharEn } else { $goreSeriesEn })
        $gtagZh = $(if ($goreCharEn) { $goreCharZh } else { $goreSeriesZh })
        if (-not $gtagZh) { $gtagZh = $gtagEn }
        $posList.Add([pscustomobject]@{ cat = '角色/作品'; id = 'character'; en = $gtagEn; zh = $gtagZh })
    }
    if ($goreTerms.Count -gt 0) {
        $grnd = New-Object System.Random
        for ($gi = 0; $gi -lt $GoreCount; $gi++) {
            $g = $goreTerms[$grnd.Next($goreTerms.Count)]
            $posList.Add([pscustomobject]@{ cat = '血腥暴力组合'; id = 'gore-combos'; en = $g.en; zh = $g.zh })
        }
    }
    else { $goreWarn = '  ⚠ 没找到血腥暴力组合库：请运行 _build\build-combos.ps1 -Src 血腥暴力组合库.xml ... 入库，或把 血腥暴力组合库.xml 放到上级目录' }
    $groupLines += '  ※ 女性向/男同（乙游·BL）系列角色：按设定只输出血腥暴力提示词，不做情色拆解'
    if ($goreWarn) { $groupLines += $goreWarn }
}

# 中文俗名：精确命中标准中文名（不做文本扩展，避免"白色"之类子串误命中；长键优先，短键若被覆盖则跳过）
# 血腥暴力模式下跳过：该模式只能出血腥暴力词，不补服装/俗名类词
if (-not $goreOnly) {
    $synMatched = New-Object System.Collections.Generic.List[string]
    # 长度降序；长度相同时保持来源顺序（确定性，避免哈希表枚举顺序导致的差异）
    $synOrdered = @()
    for ($si = 0; $si -lt $synKeys.Count; $si++) { $synOrdered += [pscustomobject]@{ k = $synKeys[$si]; i = $si } }
    $synOrdered = @($synOrdered | Sort-Object -Property @{ Expression = { $_.k.Length }; Descending = $true },
                                                        @{ Expression = { $_.i }; Descending = $false })
    foreach ($o in $synOrdered) {
        $k = $o.k
        if (-not $sLower.Contains($k)) { continue }
        $covered = $false
        foreach ($mk in $synMatched) { if ($mk.Contains($k)) { $covered = $true; break } }
        if ($covered) { continue }
        $synMatched.Add($k)
        $canon = $SYN[$k].ToLower()
        for ($i = 0; $i -lt $L_zhL.Count; $i++) {
            if ($L_zhL[$i] -eq $canon) {
                $en = $L_en[$i]
                if (-not $seen.ContainsKey($en) -and -not $cardTokens.ContainsKey($en.ToLower())) {
                    $seen[$en] = $true; $posList.Add([pscustomobject]@{ cat = $L_cat[$i]; id = $L_id[$i]; en = $en; zh = $L_zh[$i] })
                }
                break
            }
        }
    }
}

# ===== 联动词（共享常量.json 的 linkRules）：句子里命中了某类词，就补上它标志性的词 =====
# 例：出现触手 → 补「触手缠绕四肢躯干」；出现调教/束缚 → 补「高潮脸 / 阿黑颜」
$LINK = @()
if ($SC.linkRules -and $SC.linkRules.rules) { $LINK = @($SC.linkRules.rules) }
$linkNotes = New-Object System.Collections.Generic.List[string]
if (-not $goreOnly -and $LINK.Count -gt 0 -and $posList.Count -gt 0) {
    $lrand = New-Object System.Random
    $haveIds = @{}
    foreach ($p in $posList) { $haveIds[$p.id] = $true }
    foreach ($rl in $LINK) {
        $trig = $false
        foreach ($w in @($rl.when)) { if ($haveIds.ContainsKey($w)) { $trig = $true; break } }
        if (-not $trig -and $rl.whenWords) {
            $ws = @($rl.whenWords) | ForEach-Object { ([string]$_).ToLower() }
            foreach ($p in $posList) {
                if ($ws -contains $p.en.ToLower() -or $ws -contains $p.zh.ToLower()) { $trig = $true; break }
            }
        }
        if (-not $trig -and $rl.whenRx) { foreach ($p in $posList) { if ($p.en -match [string]$rl.whenRx) { $trig = $true; break } } }
        if (-not $trig) { continue }
        $cand = New-Object System.Collections.Generic.List[int]
        foreach ($c in @($rl.cats)) { if ($catIdx.ContainsKey($c)) { foreach ($ix in $catIdx[$c]) { $cand.Add($ix) } } }
        if ($cand.Count -eq 0) { continue }
        $pool = $cand
        if ($rl.prefer) {
            $prx = (@($rl.prefer) -join '|')
            $f = New-Object System.Collections.Generic.List[int]
            foreach ($ix in $cand) { if ($L_enL[$ix] -match $prx) { $f.Add($ix) } }
            if ($f.Count -gt 0) { $pool = $f }   # 优先词池非空就用它（保证"缠四肢躯干"/"高潮脸"必出）
        }
        $added = 0
        for ($k = 0; $k -lt [int]$rl.count; $k++) {
            for ($try = 0; $try -lt 14; $try++) {
                $ix = $pool[$lrand.Next($pool.Count)]
                if ($seen.ContainsKey($L_en[$ix]) -or $cardTokens.ContainsKey($L_en[$ix].ToLower())) { continue }
                $seen[$L_en[$ix]] = $true
                $posList.Add([pscustomobject]@{ cat = ('联动·' + [string]$rl.name); id = $L_id[$ix]; en = $L_en[$ix]; zh = $L_zh[$ix] })
                $added++
                break
            }
        }
        if ($added -gt 0) { $linkNotes.Add(("  ↳ 联动补词（{0}）：{1} 个" -f [string]$rl.name, $added)) }
    }
}

# 按类别分组（附加拆解部分）
$byCat = @{}
foreach ($p in $posList) { if (-not $byCat.ContainsKey($p.cat)) { $byCat[$p.cat] = New-Object System.Collections.Generic.List[string] }; $byCat[$p.cat].Add($p.en) }
foreach ($cat in $byCat.Keys) { $groupLines += ("  [{0}] {1}" -f $cat, (($byCat[$cat]) -join ' / ')) }
foreach ($ln in $linkNotes) { $groupLines += $ln }

# 换装：若输入里点名了服装（服装类词），剔除角色卡原有的服装词后再拼接
$outfit = @($posList | Where-Object { $WearIds.Contains($_.id) })
$strip  = $outfit.Count -gt 0
$cardPrompts = @($cardHit | ForEach-Object { if ($strip) { Remove-WearToken ([string]$_.prompt_en) } else { ([string]$_.prompt_en) } })
$posEn = (@($cardPrompts + @($posList | ForEach-Object { $_.en })) -join ', ')
$posZh = (@(@($cardHit | ForEach-Object { $_.prompt_zh } | Where-Object { $_ }) + @($posList | ForEach-Object { $_.zh })) -join '，')
# ===== 角色卡命中：特征明细 + 纯特征版（老模型不认识该角色时，用特征词照着画） =====
# 源用「已换装」后的文本：句子里点名了服装时，原服装已被剔除，点名服装并入「服饰」组。
if ($cardHit.Count -gt 0) {
    $fanPrefixEn = 'masterpiece, best quality, ultra-detailed, '
    $fanPrefixZh = '杰作，最佳画质，超精细，'
    $prep = @('masterpiece', 'best quality', 'ultra-detailed')

    $entryList = New-Object System.Collections.Generic.List[object]
    $need = New-Object 'System.Collections.Generic.HashSet[string]'
    for ($ci = 0; $ci -lt $cardHit.Count; $ci++) {
        $cd = $cardHit[$ci]
        $skip = @{}
        foreach ($w in $prep) { $skip[$w] = $true }
        $ne = NormFeat $cd.name_en; if ($ne) { $skip[$ne] = $true }
        $se = NormFeat $cd.series_en; if ($se) { $skip[$se] = $true }
        $toks = New-Object System.Collections.Generic.List[string]
        foreach ($tk in ([string]$cardPrompts[$ci] -split ',')) {
            $raw = $tk.Trim(); if (-not $raw) { continue }
            $k = NormFeat $raw; if (-not $k) { continue }
            [void]$need.Add($k)
            if ($skip.ContainsKey($k)) { continue }   # 质量前缀 / 角色名 / 作品名：不进特征分组
            $toks.Add($raw)
        }
        # 被换装剔除的词也要查得到中文（下面同步中文提示词要用），故按原 prompt_en 再补一遍
        foreach ($tk in ([string]$cd.prompt_en -split ',')) {
            $raw = $tk.Trim(); if (-not $raw) { continue }
            $k = NormFeat $raw; if ($k) { [void]$need.Add($k) }
        }
        $entryList.Add([pscustomobject]@{ card = $cd; toks = $toks })
    }

    # 一次性在词库数组里查出这些 token 的类别与中文（只扫一遍，命中才走这段）
    # 负向词优先判定（extra ears / long fingers 这类纠错词正是靠这条剔除）；
    # featureKeep 里的词与负向同名但确实是合法特征（如 green skin），跳过负向判定。
    $keepFeat = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($kw in $SC.featureKeep) { [void]$keepFeat.Add((NormFeat ([string]$kw))) }
    $info = @{}
    foreach ($t in $negTerms) {
        $k = NormFeat ([string]$t.en)
        if ($k -and $need.Contains($k) -and -not $keepFeat.Contains($k)) { $info[$k] = @{ id = 'negative'; zh = [string]$t.zh } }
    }
    for ($i = 0; $i -lt $L_enL.Count; $i++) {
        $k = $L_enL[$i]
        if ($k.IndexOf('_') -ge 0) { $k = $k.Replace('_', ' ') }
        if ($need.Contains($k) -and -not $info.ContainsKey($k)) { $info[$k] = @{ id = [string]$L_id[$i]; zh = [string]$L_zh[$i] } }
    }

    # 换装：卡片自带的 prompt_zh 仍含原服装词，把被剔除的英文 token 对应的中文一并去掉（只删一份、保持原序），
    # 让「正向提示词 中文」与英文严格对应。逐词表 $info 已覆盖原 prompt_en 的全部词。
    if ($strip) {
        $zhKept = New-Object System.Collections.Generic.List[string]
        for ($ci = 0; $ci -lt $cardHit.Count; $ci++) {
            $cd = $cardHit[$ci]
            $drops = New-Object System.Collections.Generic.List[string]
            foreach ($tk in ([string]$cd.prompt_en -split ',')) {
                $raw = $tk.Trim(); if (-not $raw) { continue }
                if (-not (Test-WearToken $raw)) { continue }
                $k = NormFeat $raw
                if ($info.ContainsKey($k) -and $info[$k].zh) { $drops.Add(([string]$info[$k].zh)) }
            }
            foreach ($z in ([string]$cd.prompt_zh -split '，')) {
                $zs = $z.Trim(); if (-not $zs) { continue }
                $at = -1
                for ($di = 0; $di -lt $drops.Count; $di++) { if ($drops[$di] -eq $zs) { $at = $di; break } }
                if ($at -ge 0) { $drops.RemoveAt($at); continue }
                $zhKept.Add($zs)
            }
        }
        foreach ($p in $posList) { if ($p.zh) { $zhKept.Add([string]$p.zh) } }
        $posZh = ($zhKept -join '，')
    }

    $catGroup = @{}; $groupOrder = New-Object System.Collections.Generic.List[string]
    foreach ($g in $SC.featureGroups) {
        $gn = [string]$g[0]; $groupOrder.Add($gn)
        foreach ($gid in $g[1]) { $catGroup[[string]$gid] = $gn }
    }
    $pfxList = @(); foreach ($p in $SC.featurePrefixes) { $pfxList += , @([string]$p[0], [string]$p[1]) }
    # 关键词覆盖：耳朵/尾巴/角/翅/光环 之类，按词库类别归组时会散落到"体型"或"配饰"，这里按词本身纠回一组
    $kwGroups = @(); foreach ($p in $SC.featureKw) { $kwGroups += , @([string]$p[0], @($p[1] | ForEach-Object { [string]$_ })) }
    # 服饰组名（按 top-wear 反查，跟随 共享常量.json 改名）
    $wearGrp = ''
    foreach ($g in $SC.featureGroups) { if (@($g[1]) -contains 'top-wear') { $wearGrp = [string]$g[0]; break } }

    foreach ($entry in $entryList) {
        $cd = $entry.card
        $buckets = @{}
        $keptEn = New-Object System.Collections.Generic.List[string]
        $keptZh = New-Object System.Collections.Generic.List[string]
        $negHit = New-Object System.Collections.Generic.List[string]
        $other = 0
        foreach ($raw in $entry.toks) {
            $k = NormFeat $raw
            $id = ''; $zh = ''
            if ($info.ContainsKey($k)) { $id = $info[$k].id; $zh = $info[$k].zh }
            if ($id -eq 'negative') { $negHit.Add($raw); continue }
            $grp = ''
            if ($id) {
                if ($catGroup.ContainsKey($id)) { $grp = $catGroup[$id] }
                else { foreach ($p in $pfxList) { if ($id.StartsWith($p[1])) { $grp = $p[0]; break } } }
            }
            elseif (-not $id) { $grp = '未归类' }   # 词库里查不到（约 0.7%）→ 无法判定，保留
            if (-not $grp) { $other++; continue }   # 词库里有、但不是外观特征 → 剔除
            if ($grp) {
                foreach ($kwg in $kwGroups) {
                    $hitKw = $false
                    foreach ($wt in ($k -split '[^a-z0-9]+')) { if ($wt -and ($kwg[1] -contains $wt)) { $hitKw = $true; break } }
                    if ($hitKw) { $grp = $kwg[0]; break }
                }
            }
            if (-not $buckets.ContainsKey($grp)) { $buckets[$grp] = New-Object System.Collections.Generic.List[string] }
            $buckets[$grp].Add($(if ($zh) { "$raw $zh" } else { $raw }))
            $keptEn.Add($raw)
            $keptZh.Add($(if ($zh) { $zh } else { $raw }))
        }
        # 点名服装并入「服饰」组（与上面已换装的整卡提示词保持一致）
        if ($outfit.Count -and $wearGrp) {
            if (-not $buckets.ContainsKey($wearGrp)) { $buckets[$wearGrp] = New-Object System.Collections.Generic.List[string] }
            foreach ($o in $outfit) {
                $oe = [string]$o.en; $oz = [string]$o.zh
                $buckets[$wearGrp].Add($(if ($oz) { "$oe $oz" } else { $oe }))
                $keptEn.Add($oe)
                $keptZh.Add($(if ($oz) { $oz } else { $oe }))
            }
        }

        $featLines += ("-- 角色特征 · 老模型用（不认角色名就照这个画）· {0} {1} --" -f $cd.name_zh, $cd.name_en)
        $ordered = @()
        foreach ($gn in $groupOrder) { if ($buckets.ContainsKey($gn)) { $ordered += $gn } }
        if ($buckets.ContainsKey('未归类')) { $ordered += '未归类' }
        foreach ($gn in $ordered) { $featLines += ("  [{0}] {1}" -f $gn, ($buckets[$gn] -join ' / ')) }
        if ($negHit.Count -or $other) {
            $msg = '  ⚠ 已剔除：'
            if ($negHit.Count) {
                $show = @($negHit | Select-Object -First 6) -join ', '
                $more = if ($negHit.Count -gt 6) { " 等 $($negHit.Count) 个" } else { '' }
                $msg += ("纠错/负向词 " + $show + $more)
            }
            if ($other) { $msg += $(if ($negHit.Count) { '；' } else { '' }) + ("非外观词（姿势/场景/情色/作品等）$other 个") }
            $featLines += $msg
        }
        if ($keptEn.Count) {
            $fanEn = $fanPrefixEn + ($keptEn -join ', ')
            $fanZh = $fanPrefixZh + ($keptZh -join '，')
            if (-not $strip -and ([string]$cd.features_en)) {
                # 构建时已嵌入纯特征版（features_en / features_zh）：直接用；
                # 换装时仍走现算，让被剔除的服装词同步消失
                $fanEn = $fanPrefixEn + [string]$cd.features_en
                $fanZh = $fanPrefixZh + [string]$(if ($cd.features_zh) { $cd.features_zh } else { $cd.features_en })
            }
            $featLines += ("【纯特征版 EN】" + $fanEn)
            $featLines += ("【纯特征版 中文】" + $fanZh)
        }
        else { $featLines += '  （这张卡没有可用的外观特征词）' }
    }
}
$negEn = ($negList | ForEach-Object { $_.en }) -join ', '
$negZh = ($negList | ForEach-Object { $_.zh } | Where-Object { $_ }) -join '，'

$cardHeader = if ($goreOnly) { '-- 血腥暴力提示词（女性向/男同系列规则）--' } elseif ($cardHit.Count -gt 0) { if ($strip) { "-- 角色卡命中 · 已换装 + 附加拆解 --" } else { "-- 角色卡命中 + 附加拆解 --" } } else { "-- 命中明细（按类别）--" }

$lines = @()
$lines += "【正向提示词 EN】$posEn"
$lines += "【正向提示词 中文】$posZh"
if ($strip) { $lines += "【已换装】用 " + (($outfit | ForEach-Object { $_.en }) -join ' / ') + " 替换角色卡原有服装（原服装词已移除）" }
$lines += "【反向提示词 EN】$negEn"
if ($negZh) { $lines += "【反向提示词 中文】$negZh" }
$lines += $cardHeader
$lines += ($groupLines -join "`n")
if ($featLines.Count -gt 0) { $lines += ($featLines -join "`n") }
$text = $lines -join "`n"

Write-Host $text
if ($Copy) { Set-Clipboard -Value $text; Write-Host "[已复制到剪贴板]" -ForegroundColor Green }
if ($Out) {
    $path = if ([System.IO.Path]::IsPathRooted($Out)) { $Out } else { Join-Path $root $Out }
    [System.IO.File]::WriteAllText($path, $text, (New-Object System.Text.UTF8Encoding($true)))
    Write-Host "[已写入 $path]" -ForegroundColor Green
}
