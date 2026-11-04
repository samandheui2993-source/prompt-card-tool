#Requires -Version 5.1
<#
  抽卡 - 随机抽取绘图提示词（正向 + 反向）
  数据来源: 上级目录 提示词对照.xml

  用法示例:
    .\抽卡.ps1                          # 抽 1 套完整角色提示词（含反向）
    .\抽卡.ps1 -Count 3                 # 抽 3 套
    .\抽卡.ps1 -NegCount 20             # 反向提示词抽 20 个
    .\抽卡.ps1 -All -Count 20           # 从全部词库随机抽 20 个正向词
    .\抽卡.ps1 -All -IncludeNames -Count 20   # 纯随机并包含画师/角色/作品名
    .\抽卡.ps1 -Categories hair-style,eye -Count 2
    .\抽卡.ps1 -Combos -Count 3          # 从情色组合库随机抽 3 条完整组合
    .\抽卡.ps1 -Gore -Count 3            # 从血腥暴力组合库随机抽 3 条（女性向/乙游/BL 游戏与全男系列角色用）
    .\抽卡.ps1 -Seed 12345 -Copy
    .\抽卡.ps1 -All -MinPosts 1000 -Count 20   # 热度筛选：只抽 post_count≥1000 的词（反向/组合库不受影响）
    .\抽卡.ps1 -All -WeightPosts -Count 20     # 热度加权：按 √热度 抽样，热门词更常出现；缺热度数据权重=1
    .\抽卡.ps1 -ListCategories
#>
[CmdletBinding()]
param(
    [int]$Count = 1,
    [int]$NegCount = 30,
    [switch]$All,
    [switch]$IncludeNames,
    [string[]]$Categories,
    [switch]$Combos,
    [switch]$Gore,
    [int]$MinPosts = 0,       # 热度筛选：>0 时只抽 post_count ≥ 该值的词（反向 / 组合库不受影响）
    [switch]$WeightPosts,     # 热度加权：正向词按 √热度 抽样，热门词更常出现
    [switch]$ListCategories,
    [switch]$Copy,
    [string]$Out,
    [int]$Seed
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
if ($Categories) { $Categories = @($Categories -split ',' | Where-Object { $_ -ne '' }) }
$root    = Split-Path -Parent $MyInvocation.MyCommand.Path

# ---- 共享常量（单一来源：共享常量.json；网页版由 _build\build-constants.ps1 生成 共享常量.js） ----
$scPath = Join-Path $root '共享常量.json'
if (-not (Test-Path -LiteralPath $scPath)) { $scPath = Join-Path (Split-Path -Parent $root) '共享常量.json' }
if (-not (Test-Path -LiteralPath $scPath)) { throw '找不到 共享常量.json，请放在本脚本目录或其上级目录' }
$SC = Get-Content -LiteralPath $scPath -Raw -Encoding UTF8 | ConvertFrom-Json
$nameCats = @($SC.nameCats)   # 纯随机默认排除的专有名词类别
# 黑名单词条（共享常量.json 的 wipeTerms）：名牌 / 写名字的框之类，任何模式都不进正向（负面词类别不受影响）
$wipeEn = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
if ($SC.wipeTerms) { foreach ($w in @($SC.wipeTerms.terms)) { if ($w) { [void]$wipeEn.Add(([string]$w).Trim()) } } }

$names   = @('提示词总库.xml','提示词对照.xml')
$xmlPath = $null
foreach ($n in $names) { $p = Join-Path $root $n; if (Test-Path -LiteralPath $p) { $xmlPath = $p; break } }
if (-not $xmlPath) { foreach ($n in $names) { $p = Join-Path (Split-Path -Parent $root) $n; if (Test-Path -LiteralPath $p) { $xmlPath = $p; break } } }
if (-not $xmlPath) { throw "找不到词库（提示词总库.xml 或 提示词对照.xml），请放在本脚本目录或其上级目录" }

# ---- 流式读取词库（不建 DOM）：除组合库用法外跳过组合库类别，避免 20MB+ 组合库拖慢每次抽卡 ----
$ComboCatIds = @('erotic-combos', 'gore-combos')
$wantCombos = [bool]$Combos -or [bool]$Gore -or [bool]$ListCategories
if ($Categories) { foreach ($cn2 in $Categories) { if ($ComboCatIds -contains $cn2 -or $cn2 -eq '情色组合' -or $cn2 -eq '血腥暴力组合') { $wantCombos = $true } } }

$db    = @{}
$order = New-Object System.Collections.Generic.List[string]

function Read-Library([string]$path, [bool]$withCombos) {
    $rs = New-Object System.Xml.XmlReaderSettings
    $rs.IgnoreComments = $true
    $rs.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
    $rd = [System.Xml.XmlReader]::Create($path, $rs)
    try {
        $cur = $null; $skip = $false
        while ($rd.Read()) {
            if ($rd.NodeType -eq [System.Xml.XmlNodeType]::Element) {
                if ($rd.Name -eq 'category') {
                    $cid = $rd.GetAttribute('id'); $czh = $rd.GetAttribute('zh')
                    if ($ComboCatIds -contains $cid -and -not $withCombos) { $skip = $true; $cur = $null }
                    else {
                        $skip = $false
                        # 每个词只存平行的 en / tzh / posts 三个数组，不建逐词对象：解析约快 2 倍
                        $cur = [pscustomobject]@{ id = $cid; zh = $czh
                                                 en  = (New-Object 'System.Collections.Generic.List[string]')
                                                 tzh = (New-Object 'System.Collections.Generic.List[string]')
                                                 posts = (New-Object 'System.Collections.Generic.List[int]') }
                        $db[$cid] = $cur; $order.Add($cid)
                    }
                }
                elseif (-not $skip -and $rd.Name -eq 'term' -and $cur) {
                    $enT = [string]$rd.GetAttribute('en')
                    if ($cur.id -ne 'negative' -and $wipeEn.Count -gt 0 -and $wipeEn.Contains($enT.Trim())) { continue }   # 黑名单词条不进正向
                    $pv = 0; [void][int]::TryParse([string]$rd.GetAttribute('posts'), [ref]$pv)
                    $cur.en.Add($enT); $cur.tzh.Add($rd.GetAttribute('zh')); $cur.posts.Add($pv)
                }
            }
            elseif ($rd.NodeType -eq [System.Xml.XmlNodeType]::EndElement -and $rd.Name -eq 'category') {
                $skip = $false; $cur = $null
            }
        }
    }
    finally { $rd.Close() }
}

Read-Library $xmlPath $wantCombos
$hasNeg = $db.ContainsKey('negative')

# 组合库：若词库本身不含（例如只找到 提示词对照.xml），尝试单独加载组合库文件
if ($wantCombos) {
    if (-not $db.ContainsKey('erotic-combos')) {
        foreach ($cn in @('情色提示词组合库.xml')) {
            $cp = Join-Path $root $cn
            if (-not (Test-Path -LiteralPath $cp)) { $cp = Join-Path (Split-Path -Parent $root) $cn }
            if (Test-Path -LiteralPath $cp) { Read-Library $cp $true; break }
        }
    }
    if (-not $db.ContainsKey('gore-combos')) {
        foreach ($cn in @('血腥暴力组合库.xml')) {
            $cp = Join-Path $root $cn
            if (-not (Test-Path -LiteralPath $cp)) { $cp = Join-Path (Split-Path -Parent $root) $cn }
            if (Test-Path -LiteralPath $cp) { Read-Library $cp $true; break }
        }
    }
}

if ($ListCategories) {
    $order | ForEach-Object { '{0,-26} {1,-16} {2}词' -f $_, $db[$_].zh, $db[$_].en.Count }
    return
}

$rand = if ($Seed) { New-Object System.Random($Seed) } else { New-Object System.Random }

function Resolve-CategoryId([string]$name) {
    if ($db.ContainsKey($name)) { return $name }
    foreach ($k in $db.Keys) { if ($db[$k].zh -eq $name) { return $k } }
    throw "未知类别: $name (用 -ListCategories 查看)"
}
# ---- 热度池（posts 属性；缺热度=0）：-MinPosts 过滤、-WeightPosts 按 √热度 加权 ----
$poolCache = @{}
function Get-Pool([string]$id) {
    # 返回 $null = 无热度数据（组合库等）：调用方均匀回退；
    # 返回 @{idx=@()} 空池 = 启用了 -MinPosts 但整类不达标：Draw-All 跳过该类 / New-Pick 均匀回退；
    # 否则 { idx, cum(加权前缀和或 $null), total }
    if (-not (($MinPosts -gt 0) -or $WeightPosts)) { return $null }
    $ck = "${id}|$MinPosts|$([bool]$WeightPosts)"
    if ($poolCache.ContainsKey($ck)) { return $poolCache[$ck] }
    $p = $null
    $c = $db[$id]
    if ($c -and $c.posts.Count -eq $c.en.Count) {
        $idx = New-Object 'System.Collections.Generic.List[int]'
        if ($MinPosts -gt 0) {
            for ($i = 0; $i -lt $c.en.Count; $i++) { if ($c.posts[$i] -ge $MinPosts) { $idx.Add($i) } }
        }
        else { for ($i = 0; $i -lt $c.en.Count; $i++) { $idx.Add($i) } }
        if ($idx.Count -gt 0) {
            if (-not $WeightPosts) { $p = @{ idx = $idx; cum = $null; total = 0.0 } }
            else {
                $cum = New-Object 'double[]' $idx.Count
                $t = 0.0
                for ($i = 0; $i -lt $idx.Count; $i++) {
                    $w = [double]$c.posts[$idx[$i]]; if ($w -lt 1) { $w = 1.0 }
                    $t += [Math]::Sqrt($w); $cum[$i] = $t
                }
                $p = @{ idx = $idx; cum = $cum; total = $t }
            }
        }
        else {
            if ($MinPosts -gt 0) { $p = @{ idx = $idx; cum = $null; total = 0.0 } }   # 空池哨兵（整类低于阈值）
            # 权重态（无阈值）时 $idx 为全量，不会走到这里；空池哨兵仅在 -MinPosts 开启时出现
        }
    }
    $poolCache[$ck] = $p
    return $p
}
function New-Pick([string]$id) {
    $c = $db[$id]
    $pool = Get-Pool $id
    $i = -1
    if ($pool) {
        if ($pool.cum) {
            $x = $rand.NextDouble() * $pool.total
            $lo = 0; $hi = $pool.cum.Length - 1
            while ($lo -lt $hi) { $mid = ($lo + $hi) -shr 1; if ($x -lt $pool.cum[$mid]) { $hi = $mid } else { $lo = $mid + 1 } }
            $i = $pool.idx[$lo]
        }
        else { $i = $pool.idx[$rand.Next($pool.idx.Count)] }
    }
    if ($i -lt 0) { $i = $rand.Next($c.en.Count) }   # 未启用 / 池空（该类没有达标词）→ 均匀回退
    return [pscustomobject]@{ id = $id; catZh = $c.zh; en = $c.en[$i]; zh = $c.tzh[$i] }
}
function Draw-Neg([int]$n) {
    if (-not $hasNeg) { return @() }
    $neg = $db['negative']; $cnt = $neg.en.Count
    $ord = New-Object 'int[]' $cnt
    for ($i = 0; $i -lt $cnt; $i++) { $ord[$i] = $i }
    for ($i = $cnt - 1; $i -gt 0; $i--) { $j = $rand.Next($i + 1); $tmp = $ord[$i]; $ord[$i] = $ord[$j]; $ord[$j] = $tmp }
    $take = [Math]::Min($n, $cnt)
    $out = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $take; $i++) { $k = $ord[$i]; $out.Add([pscustomobject]@{ id = 'negative'; catZh = $neg.zh; en = $neg.en[$k]; zh = $neg.tzh[$k] }) }
    return $out
}

# ---- 完整角色配方（来源: 共享常量.json） ----
$recipeBase = @($SC.recipeBase)
$recipePool = @($SC.recipePool) |
                Where-Object { $db.ContainsKey($_) }

function Draw-Set {
    $pos = New-Object System.Collections.Generic.List[object]
    foreach ($id in $recipeBase) { if ($db.ContainsKey($id)) { $pos.Add((New-Pick $id)) } }
    $extra = $recipePool | Sort-Object { $rand.Next() } | Select-Object -First ($rand.Next(2, 5))
    foreach ($id in $extra) { $pos.Add((New-Pick $id)) }
    return [pscustomobject]@{ pos = $pos; neg = (Draw-Neg $NegCount) }
}
function Draw-Cats([string[]]$ids) {
    $pos = New-Object System.Collections.Generic.List[object]
    foreach ($name in $ids) { $id = Resolve-CategoryId $name; if ($id -ne 'negative') { $pos.Add((New-Pick $id)) } }
    return [pscustomobject]@{ pos = $pos; neg = (Draw-Neg $NegCount) }
}
function Draw-All([int]$n) {
    $skip = @('negative') + $ComboCatIds; if (-not $IncludeNames) { $skip += $nameCats }
    if (($MinPosts -gt 0) -or $WeightPosts) {
        # 两步抽样：类别按候选数占比 → 类内按热度规则（与全词扁平均匀同分布）
        $catIds = New-Object System.Collections.Generic.List[string]
        $cum = New-Object System.Collections.Generic.List[double]
        $total = 0.0
        foreach ($id in $order) {
            if ($skip -contains $id) { continue }
            $c = $db[$id]; if ($c.en.Count -eq 0) { continue }
            $pool = Get-Pool $id
            $cnt = 0
            if ($pool) { if ($pool.idx.Count -eq 0) { continue }; $cnt = $pool.idx.Count }
            else { $cnt = $c.en.Count }
            $total += $cnt
            $cum.Add($total)
            $catIds.Add($id)
        }
        $pos = New-Object System.Collections.Generic.List[object]
        for ($i = 0; $i -lt $n; $i++) {
            $x = $rand.NextDouble() * $total
            $k = 0
            while ($cum[$k] -lt $x) { $k++ }        # $x < $total，不越界
            $pos.Add((New-Pick $catIds[$k]))
        }
        return [pscustomobject]@{ pos = $pos; neg = (Draw-Neg $NegCount) }
    }
    # 只累计各类别的起始下标，抽样时按全局下标定位，避免展开 16 万条的扁平列表
    $ranges = New-Object System.Collections.Generic.List[object]
    $total = 0
    foreach ($id in $order) {
        if ($skip -contains $id) { continue }
        $c = $db[$id]; if ($c.en.Count -eq 0) { continue }
        $ranges.Add([pscustomobject]@{ id = $id; catZh = $c.zh; start = $total; count = $c.en.Count })
        $total += $c.en.Count
    }
    $pos = New-Object System.Collections.Generic.List[object]
    if ($total -gt 0) {
        for ($i = 0; $i -lt $n; $i++) {
            $k = $rand.Next($total)
            foreach ($r in $ranges) {
                if ($k -lt $r.start + $r.count) {
                    $c = $db[$r.id]; $j = $k - $r.start
                    $pos.Add([pscustomobject]@{ id = $r.id; catZh = $r.catZh; en = $c.en[$j]; zh = $c.tzh[$j] })
                    break
                }
            }
        }
    }
    return [pscustomobject]@{ pos = $pos; neg = (Draw-Neg $NegCount) }
}

function Join-En($picks) { ($picks | ForEach-Object { $_.en }) -join ', ' }
function Join-Zh($picks) { ($picks | ForEach-Object { $_.zh }) -join '，' }
function Detail($picks)   { ($picks | ForEach-Object { '  {0,-24} {1} ({2})' -f $_.catZh, $_.en, $_.zh }) -join "`n" }

function Format-Sets($sets) {
    $blocks = @(); $idx = 0
    foreach ($s in $sets) {
        $idx++
        if ($sets.Count -gt 1) { $blocks += "━━━ 第 $idx 套 ━━━" }
        $blocks += "【正向提示词 EN】" + (Join-En $s.pos)
        $blocks += "【正向提示词 中文】" + (Join-Zh $s.pos)
        $blocks += "【反向提示词 EN】" + (Join-En $s.neg)
        $blocks += "【反向提示词 中文】" + (Join-Zh $s.neg)
        if (-not $Combos -and -not $Gore) {
            $blocks += "-- 正向明细 --"
            $blocks += (Detail $s.pos)
        }
        $blocks += ""
    }
    return ($blocks -join "`n")
}

function Draw-Combos([string]$id, [int]$n) {
    $rc = $db[$id]
    if (-not $rc -or $rc.en.Count -eq 0) { throw ("组合库为空（{0}）：请先运行 _build\build-combos.ps1，或把对应组合库 xml 放到上级目录" -f $id) }
    # 一次性洗牌组合库索引，按需取前 n 个（n 超过库容量时循环取）
    $ord = New-Object 'int[]' $rc.en.Count
    for ($i = 0; $i -lt $ord.Length; $i++) { $ord[$i] = $i }
    for ($i = $ord.Length - 1; $i -gt 0; $i--) { $j = $rand.Next($i + 1); $tmp = $ord[$i]; $ord[$i] = $ord[$j]; $ord[$j] = $tmp }
    $out = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $n; $i++) {
        $k = $ord[$i % $ord.Length]
        $out.Add([pscustomobject]@{ pos = @([pscustomobject]@{ id = $id; catZh = $rc.zh; en = $rc.en[$k]; zh = $rc.tzh[$k] }); neg = (Draw-Neg $NegCount) })
    }
    return $out
}

$sets = @()
if ($All) { $sets = ,(Draw-All $Count) }
elseif ($Categories) { for ($i = 0; $i -lt $Count; $i++) { $sets += ,(Draw-Cats $Categories) } }
elseif ($Gore) { $sets = @(Draw-Combos 'gore-combos' $Count) }
elseif ($Combos) { $sets = @(Draw-Combos 'erotic-combos' $Count) }
else { for ($i = 0; $i -lt $Count; $i++) { $sets += ,(Draw-Set) } }

$text = Format-Sets $sets
Write-Host $text
if ($Copy) { Set-Clipboard -Value $text; Write-Host "[已复制到剪贴板]" -ForegroundColor Green }
if ($Out) {
    $path = if ([System.IO.Path]::IsPathRooted($Out)) { $Out } else { Join-Path $root $Out }
    [System.IO.File]::WriteAllText($path, $text, (New-Object System.Text.UTF8Encoding($true)))
    Write-Host "[已写入 $path]" -ForegroundColor Green
}
