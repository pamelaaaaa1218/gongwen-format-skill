# 国企体制内 Word 公文排版技能

把 Word 文档（.docx / .doc）一键排版成国企标准化公文格式的 Claude Code 技能。
适用于工作汇报、报告、通知、请示、总结等公文类文档。

## 在新电脑上安装

1. 克隆到 Claude Code 的用户技能目录：

   ```
   gh repo clone pamelaaaaa1218/gongwen-format-skill ~/.claude/skills/gongwen-format
   ```

   没装 gh 时用 git（私有库会提示登录）：

   ```
   git clone https://github.com/pamelaaaaa1218/gongwen-format-skill.git ~/.claude/skills/gongwen-format
   ```

2. 安装依赖：

   ```
   pip3 install python-docx
   ```

3. 重新打开 Claude Code，技能即可使用。

## 用法

对 Claude Code 说「把这个文档套成公文格式」并给出文件路径，或输入 `/gongwen-format`。
`.docx` 与 `.doc` 均支持（`.doc` 在 macOS 上会自动转换；其他系统请先在 Word/WPS 里另存为 `.docx`）。

## 格式规格

| 段落 | 字体 | 字号 |
|------|------|------|
| 标题 / 单位抬头 | 方正小标宋简体 | 小二 18pt，居中 |
| 正文 | 仿宋_GB2312 | 三号 16pt，首行缩进 2 字 |
| 一级标题「一、」 | 黑体 | 三号 16pt，段前 1 行 |
| 二级标题「（一）」 | 楷体_GB2312 | 三号 16pt，段前 0.5 行 |
| 三级「1.」/ 四级「（1）」/ 附件 / 落款 | 仿宋_GB2312 | 三号 16pt |

数字与英文一律 Times New Roman（中文才用上表字体）；表格整体段前 0.5 行。
行距固定 28 磅，标题段后 1.5 行，A4 纸，页边距上 3.7 / 下 3.5 / 左 2.8 / 右 2.6 cm。

自动规范化（均默认开启，可用 `--keep-numbering / --keep-quotes / --keep-spaces` 关闭）：
二级编号「1.1」→「（一）」；无三级序号时「（1）」提升为「1.」；英文双引号 → 中文「“ ”」；
清理数字/英文两侧多余空格（保留 Claude Code、GB/T 9704 这类必要空格）。

完整规格与调整方法见 [`references/format_spec.md`](references/format_spec.md)。
