---
name: visio-sci-figure
description: 按统一的科研论文插图风格，用 Visio COM 自动绘制可编辑的 .vsdx 工程并导出高分辨率 PNG（系统框图、硬件架构图、模块结构图、信号流程图等）。风格：加粗 Times New Roman + 宋体、2~2.5 pt 粗线、灰阶填充、彩色粗框分区且区名置于框下、圆角虚线分组、实心箭头。当用户要求画论文/报告插图、框图、架构图、"画图1.x"、"做成visio"或"按我的绘图风格画"时使用。需要 Windows + 已安装 Visio + pywin32。
---

# Visio 科研插图风格

该风格从用户手绘的论文 Visio 工程中总结而来，目标是插入 Word 后线条清晰、文字醒目，黑白打印也能分辨。

## 风格规范

**字体**
- 英文、数字：Times New Roman；中文：宋体。
- 所有文字加粗（Char.Style = 17）。
- 变量、物理量符号用加粗斜体（Char.Style = 51），带上下标，例如 *U*ⱼ、*θ*ⱼ、*w*ⱼᵢ、lane×*L*。

**尺寸**（先放大绘制，插入 Word 时整体缩小）
- 画布宽度一般为 250~400 mm，插入 Word 时缩到版心宽度（约 15 cm）。
- 正文字号 13~14 pt，分区名和标题 18 pt，次要注释 11 pt。缩小后约为 6~8 pt。
- 线宽：普通连线和模块边框 2 pt，分区框 2.5 pt，高速总线或主数据通路 4~4.5 pt。

**模块框**
- 黑色 2 pt 边框，2 mm 小圆角，不加阴影。
- 填充以灰阶为主：普通模块用浅灰 242，核心模块用中灰 216。
- 存储器类模块用深灰填充，配白色加粗字（`memory()`）。
- 少量强调色：控制器或关键模块用浅粉 (242,220,218)，另有浅蓝 (219,238,243)、浅绿 (235,241,222)、淡黄。同一张图里强调色不超过 2~3 种。

**分区与分组**
- 功能分区（芯片、子系统、处理核）用蓝色 (0,0,255) 或红色 (255,0,0) 的 2.5 pt 直角粗框。
- 分区名用同色加粗 18 pt 字，放在框的正下方（`region()`）。
- 子模块组用彩色圆角虚线框包围（LinePattern 23，`group()`），名称写在框上方。

**连线与箭头**
- 正交折线，箭头为实心三角，中等大小。
- 不同信号用颜色加线型区分，信号名用与线同色的加粗字。示例：时钟用黑色实线，参考或同步信号用紫色虚线，反馈信号用绿色虚线。
- 线色避免与分区框颜色相同，免得混淆。
- 最终输出（如"识别结果"）用红色粗块状箭头配红字（`block_arrow()`）。
- 外部输入、输出从图的左右两侧水平进出，下方留给分区名。

**其他**
- 省略号用粗黑点"⋮""…"。
- 布局保持对称、对齐，并留足间距，标签不要压线。

## 工作流程

1. **规划布局**：先在 mm 坐标系（左上角为原点）中确定分区、模块、连线和标签的位置。按文字宽度估算模块大小：14 pt 加粗时，中文约 5 mm/字，英文大写和数字约 3 mm/字。
2. **编写脚本**：用 `scripts/vsd_style.py` 写绘图脚本，放在用户项目的 figures 目录下，例如 `figures/fig1_1.py`。脚本开头先把本 skill 的 scripts 目录加入 `sys.path`：
   ```python
   import sys, os
   sys.path.insert(0, os.path.expanduser('~/.agents/skills/visio-sci-figure/scripts'))
   from vsd_style import Figure, C, BOLD_ITALIC

   with Figure(352, 128, os.path.join(os.path.dirname(__file__), 'fig1_1')) as f:
       f.region(28, 76, 98, 114, 'ADC', color=C.BLUE)
       f.box(32, 82, 54, 108, 'ADC\n内核', fill=C.GRAY1)
       f.box(64, 82, 94, 108, 'JESD204B\n发送端', fill=C.GRAY2)
       f.arrow([(54, 95), (64, 95)])
       lab = f.label(111, 84, 'lane×L', 13)
       f.char_style(lab, 5, 6, BOLD_ITALIC)          # L 加粗斜体
   ```
   完整示例见 `examples/jesd204b_system.py`，即 JESD204B 系统框图。
3. **运行检查**：运行脚本，会生成 `.vsdx` 和 300 dpi 的 `.png`。用 Read 查看 PNG，检查文字溢出、压线、重叠和对齐。有问题就调整坐标后重跑，直到成图干净为止。
4. **不覆盖用户文件**：输出前先确认目标文件不存在。若已存在（用户可能手动改过），改用 `_v2` 等新文件名。
5. **插入 Word**：用 python-docx 按版心宽度插入 PNG，图片段落居中、无首行缩进、单倍行距。Word 打开文档时写入会报 PermissionError，需要提示用户先关闭文档。

## API 速查（`scripts/vsd_style.py`）

| 方法 | 用途 |
|---|---|
| `Figure(w_mm, h_mm, out_base, dpi=300)` | 创建画布；退出 `with` 时保存 `out_base.vsdx` 和 `out_base.png` |
| `box(x0,y0,x1,y1, txt, fill, line, weight, pattern, round_mm, size, color)` | 模块框 |
| `label(x, y, txt, size, color, align, w, h)` | 无边框文字；align 可取 center/left/right |
| `arrow(pts, color, weight, pattern, head, head_size)` | 折线箭头；pattern 可取 solid/dash/dashdot/round_dash；粗线时用 head_size=0 |
| `line(pts, ...)` | 无箭头连线 |
| `region(x0,y0,x1,y1, name, color)` | 彩色粗框分区，区名置于框下 |
| `group(x0,y0,x1,y1, color, name)` | 圆角虚线分组框 |
| `memory(x0,y0,x1,y1, txt)` | 深灰底白字存储器模块 |
| `block_arrow(x, y, direction, length, width, color)` | 红色块状输出箭头 |
| `Figure.char_style(sh, begin, end, style, sub, sup)` | 设置文字片段格式：加粗斜体、上下标 |

配色常量都在 `C` 中，包括 BLACK、BLUE、RED、PURPLE、GREEN、GRAY1/2/3、DARK、PINK、LBLUE、LGREEN、LYELLOW。

## 注意事项

- Visio COM 在后台运行（`Visio.InvisibleApp`）。脚本如果异常中断，可能残留 VISIO.EXE 进程，可以用 `tasklist | grep -i visio` 检查。
- 读取用户已有的大型 .vsdx 时，用 `Documents.OpenEx(path, 2+8)` 以只读方式打开，并设置超时。
- 环境里的 matplotlib 可能与 numpy 2.x 不兼容，不要依赖它出图。
