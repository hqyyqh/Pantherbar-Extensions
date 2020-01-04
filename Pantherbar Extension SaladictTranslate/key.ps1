# Pantherbar 传入的参数
param(
    [string]$TextProcess, # 处理文本判断
    [string]$PLAIN_TEXT     # 纯文本
)

<#
SaladictTranslation extension for Pantherbar——By Pencilq
该 Powershell 脚本用于文本处理以及快捷键调用
#>

# 自定义快捷键，自定义快捷键说明 https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.sendkeys?view=netframework-4.8#remarks
# 比如 Alt+l 为 "%l" ; Ctrl+l 为 "^l" ; Shift+l 为 "+l" ; Alt+Shift+l 为 "%+l"
$ShortCut = "%l"

# 文本段落正则匹配
# 末尾结束标点，与下一句分段
$RegexEndPunctuation = '[\\.:;!。！？：\s]$'
# 开头项目标点，与前一句分段
$RegexItemIdentify = '^•|^–\s|^Chapter\s[1-9][0-9]{0,1}|^(\d*\.)+\d*\s|^\d*\.\s|^·|^\[\d*\]\s[A-Z]'
# 末尾常见缩写词，防误判分段
$RegexEndAbbr = ' fig\.$| et al\.$| Fig\.$| Eq\.$| eq\.$| p\.$| pp\.$| Ph\.D\.$|cf\.$|Cf\.$|,\s\d{4};$|\.\s\(\d{4}\);$'
# 末尾英语字母或数字
$RegexEndEng = '[a-z0-9]$'

# 测试变量

<############
$TextProcess = "process"
$PLAIN_TEXT = ""
###########>

####################################################

$wshell = New-Object -ComObject wscript.shell

# 文本的分割和清洗
function TextSplit ($InputText) {       
    $InputText = $InputText -split "`n"     # 按段落分割
    for ($i = 0; $i -lt $InputText.Count; $i++) {
        $InputText[$i] = $InputText[$i].Replace("ﬁ", "fi")      # LaTeX 中连字符替换
        $InputText[$i] = $InputText[$i].Replace("ﬃ", "ffi")
        $InputText[$i] = $InputText[$i].Replace("ﬂ", "fl")
        $InputText[$i] = $InputText[$i].Replace("ﬀ", "ff")
        $InputText[$i] = $InputText[$i].Trim()                   # 删除字符串前后分行和空白字符
    }
    $InputText = $InputText.where( { $_ -ne "" })                # 删除空白行
    return $InputText
}

# 匹配标点，记录分段信息
# $Separator 设置三种情况，0 为中文直接拼接，1 为英文拼接加空格，2 为换行分段
function PunctuationSegment ($InputText){
    $Separator = @()
    for ($i = 0; $i -lt $InputText.Count-1; $i++) {
        $IsEnd = $InputText[$i] -cmatch $RegexEndPunctuation                                             # 断句符判断
        $IsNextItem = $InputText[$i + 1] -cmatch $RegexItemIdentify                                      # 列表项判断
        $IsAbbr = $InputText[$i] -cmatch $RegexEndAbbr                                                   # 结尾缩写项判断
        $IsEng = $InputText[$i] -cmatch $RegexEndEng                                                     # 英语及数字判断
        # 当前句末尾以及下一句开头是否含有分段标识符
        if ($IsEnd -or $IsNextItem) {
        # 排除末尾缩写的情况
            if ($IsAbbr) {
                $Separator += 1
            }
            else {
                $Separator += 2
            }
        }
        elseif ($IsEng) {
            $Separator += 1
        }
        else {
            $Separator += 0
        }
    }
    return $Separator
}

# 根据数组信息拼接文本
# $Separator 设置三种情况，0 为中文直接拼接，1 为英文拼接加空格，2 为换行分段
function Textjoin ($InputText, $Separator) {
    if ($InputText.Count -le 1) {
        # 只有一行字符无需处理
        $CombText = $InputText
    }
    else {
        $CombText = $InputText[0]
        for ($i = 0; $i -lt $InputText.Count - 1; $i++) {
            $CombTextSplit = $CombText -split "`n"
            $IsTitle = $CombTextSplit[-1] -cmatch $RegexItemIdentify -and $InputText[$i + 1] -cmatch '^[A-Z]'    # 标题判断
            if ($IsTitle) {
                $CombText = $CombText + "`n`n" + $InputText[$i + 1]
            }
            elseif ($Separator[$i] -eq 0) {
                $CombText = $CombText + $InputText[$i + 1]  # 中文拼接
            }
            elseif ($Separator[$i] -eq 1) {
                $CombText = $CombText + " " + $InputText[$i + 1]   # 加空格
            }
            elseif ($Separator[$i] -eq 2) {
                $CombText = $CombText + "`n`n" + $InputText[$i + 1]  # 分段
            }
        }
    }
    return $CombText
}

# 定义处理文本模块
function TextModify () {
    $SepText = TextSplit($PLAIN_TEXT)
    $Separator = PunctuationSegment($SepText)
    $JoinText = Textjoin $SepText $Separator
    $JoinText | Set-Clipboard
    $wshell.SendKeys($ShortCut)
}

# 定义直接赋值文本模块
function TextCopy () {
    $PLAIN_TEXT | Set-Clipboard
    $wshell.SendKeys($ShortCut)
}

if ($TextProcess -eq "process") {
    TextModify
}
elseif ($TextProcess -eq "initial") {
    TextCopy
}

