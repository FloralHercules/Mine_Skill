# Mine Skill

两个用于可编辑科研绘图的 Codex 技能，各自放在独立目录中。

| 目录 | 用途 | 输出 |
| --- | --- | --- |
| [image-to-visio](image-to-visio/SKILL.md) | 根据图片或截图重建框架图，使用原生形状、文字和 MathType 公式 | 可编辑 Visio 页面 |
| [visio-to-editable-ppt](visio-to-editable-ppt/SKILL.md) | 将指定 Visio 页面转换为 PowerPoint，保留原生形状、可编辑文字、独立图片和 MathType 对象 | 单页可编辑 PPTX |

## 安装

将需要的技能目录复制到 Codex 的技能目录 `$CODEX_HOME/skills/`；未设置 `CODEX_HOME` 时，通常为 `~/.codex/skills/`。两个技能可以分别安装。

这两个技能面向 Windows 桌面环境。依赖需要在本机安装，仓库不包含 Office、MathType、字体或 Python 运行库。

- **image-to-visio**：Microsoft Visio、Windows Python、`pywin32` 和 `Pillow`。生成 MathType 公式还需要 Microsoft Word、已注册的 MathType 和可用的 TeX Toggle。具体配置见 [运行环境说明](image-to-visio/references/runtime.md)。
- **visio-to-editable-ppt**：Microsoft Visio、Microsoft PowerPoint、PowerShell、Python 3.10+、`lxml` 和 `Pillow`。编辑保留的 MathType 对象需要兼容的公式编辑器。

## 使用示例

- `用 $image-to-visio 将这张框架图重建为可编辑 Visio 页面，并加入当前项目的主文件。`
- `用 $visio-to-editable-ppt 将这个 VSDX 的指定页面转成一页可编辑 PPT。`

转换后需检查预览图。图片重建需要人工判断图中语义；复杂三维投影等内容可能需要额外处理。详细步骤、限制和检查方式见各目录中的 `SKILL.md`。
