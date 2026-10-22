"""科研插图Visio绘图库：用Visio COM按统一风格生成可编辑的.vsdx与高分辨率PNG。

风格要点（详见SKILL.md）：加粗Times New Roman+宋体、2~2.5 pt粗线、灰阶填充、
彩色粗框分区且区名置于框下、圆角虚线分组、实心箭头。
坐标单位mm，左上角为原点。

用法：
    from vsd_style import Figure, C
    with Figure(352, 128, 'out/fig1_1') as f:      # 输出 out/fig1_1.vsdx 与 out/fig1_1.png
        f.box(10, 10, 40, 30, 'ADC\\n内核', fill=C.GRAY1)
        f.arrow([(40, 20), (60, 20)])
"""
import os
import win32com.client


class C:
    """配色（与用户手绘图一致）。"""
    BLACK = 'RGB(0,0,0)'
    WHITE = 'RGB(255,255,255)'
    BLUE = 'RGB(0,0,255)'          # 分区框/强调
    RED = 'RGB(255,0,0)'           # 分区框/强调/输出
    PURPLE = 'RGB(112,48,160)'
    GREEN = 'RGB(0,128,0)'
    GRAY1 = 'RGB(242,242,242)'     # 浅灰：普通模块
    GRAY2 = 'RGB(216,216,216)'     # 中灰：核心模块
    GRAY3 = 'RGB(191,191,191)'     # 深一档灰
    DARK = 'RGB(89,89,89)'         # 深灰：存储器类模块（配白字）
    PINK = 'RGB(242,220,218)'      # 浅粉：控制器/关键模块
    LBLUE = 'RGB(219,238,243)'
    LGREEN = 'RGB(235,241,222)'
    LYELLOW = 'RGB(255,242,204)'


# Visio单元格常量
_PAT = {'solid': 1, 'dash': 2, 'dashdot': 4, 'round_dash': 23, 'none': 0}
_ALIGN = {'center': 1, 'left': 0, 'right': 2}
BOLD, BOLD_ITALIC = 17, 51         # Char.Style：加粗 / 加粗斜体


class Figure:
    def __init__(self, width_mm, height_mm, out_base, dpi=300, font_en='Times New Roman', font_cn='宋体'):
        self.W, self.H = width_mm, height_mm
        self.out_base = os.path.abspath(out_base)
        self.dpi = dpi
        self.app = win32com.client.Dispatch('Visio.InvisibleApp')
        self.app.AlertResponse = 7
        self.doc = self.app.Documents.Add('')
        self.page = self.app.ActivePage
        ps = self.page.PageSheet
        ps.CellsU('PageWidth').FormulaU = '%s mm' % width_mm
        ps.CellsU('PageHeight').FormulaU = '%s mm' % height_mm
        self.f_en = self.doc.Fonts.Item(font_en).ID
        self.f_cn = self.doc.Fonts.Item(font_cn).ID

    # ---- 上下文管理：退出时保存并关闭Visio ----
    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        try:
            if exc_type is None:
                self.save()
        finally:
            self.doc.Saved = True
            self.doc.Close()
            self.app.Quit()

    def save(self):
        os.makedirs(os.path.dirname(self.out_base), exist_ok=True)
        self.doc.SaveAs(self.out_base + '.vsdx')
        self.app.Settings.SetRasterExportResolution(3, self.dpi, self.dpi, 0)
        self.page.Export(self.out_base + '.png')

    # ---- 坐标换算：mm（左上原点）→ inch（左下原点） ----
    def _x(self, mm):
        return mm / 25.4

    def _y(self, mm):
        return (self.H - mm) / 25.4

    def _text(self, sh, txt, size, color, style, align='center'):
        sh.Text = txt
        sh.CellsU('Char.Font').ResultIU = self.f_en
        sh.CellsU('Char.AsianFont').ResultIU = self.f_cn
        sh.CellsU('Char.Size').FormulaU = '%s pt' % size
        sh.CellsU('Char.Style').FormulaU = str(style)
        sh.CellsU('Char.Color').FormulaU = color
        sh.CellsU('Para.SpLine').FormulaU = '-110%'
        sh.CellsU('Para.HorzAlign').ResultIU = _ALIGN[align]
        for c in ('TopMargin', 'BottomMargin', 'LeftMargin', 'RightMargin'):
            sh.CellsU(c).FormulaU = '0 mm'

    def _rect(self, x0, y0, x1, y1):
        return self.page.DrawRectangle(self._x(x0), self._y(y1), self._x(x1), self._y(y0))

    # ---- 基本元素 ----
    def box(self, x0, y0, x1, y1, txt='', fill=None, line=C.BLACK, weight=2, pattern='solid',
            round_mm=2, size=14, color=C.BLACK, style=BOLD):
        """模块框。fill=None为无填充；深色填充时请设color=C.WHITE。"""
        sh = self._rect(x0, y0, x1, y1)
        if fill:
            sh.CellsU('FillPattern').ResultIU = 1
            sh.CellsU('FillForegnd').FormulaU = fill
        else:
            sh.CellsU('FillPattern').ResultIU = 0
        sh.CellsU('LineColor').FormulaU = line
        sh.CellsU('LineWeight').FormulaU = '%s pt' % weight
        sh.CellsU('LinePattern').ResultIU = _PAT[pattern]
        sh.CellsU('Rounding').FormulaU = '%s mm' % round_mm
        sh.CellsU('ShdwPattern').ResultIU = 0
        self._text(sh, txt, size, color, style)
        return sh

    def label(self, x, y, txt, size=14, color=C.BLACK, align='center', w=40, h=8, style=BOLD):
        """无边框文字。align=center时x为中心；left/right时x为文字起点/终点。"""
        x0 = {'center': x - w / 2, 'left': x, 'right': x - w}[align]
        sh = self._rect(x0, y - h / 2, x0 + w, y + h / 2)
        sh.CellsU('FillPattern').ResultIU = 0
        sh.CellsU('LinePattern').ResultIU = 0
        self._text(sh, txt, size, color, style, align)
        return sh

    def arrow(self, pts, color=C.BLACK, weight=2, pattern='solid', head='end', head_size=1):
        """正交折线箭头，pts为[(x,y),...]。head: end/begin/both/none；粗线(≥4 pt)建议head_size=0。"""
        xy = []
        for x, y in pts:
            xy += [self._x(x), self._y(y)]
        sh = self.page.DrawPolyline(xy, 0) if len(pts) > 2 else self.page.DrawLine(*xy)
        sh.CellsU('LineColor').FormulaU = color
        sh.CellsU('LineWeight').FormulaU = '%s pt' % weight
        sh.CellsU('LinePattern').ResultIU = _PAT[pattern]
        if head in ('end', 'both'):
            sh.CellsU('EndArrow').ResultIU = 4
            sh.CellsU('EndArrowSize').ResultIU = head_size
        if head in ('begin', 'both'):
            sh.CellsU('BeginArrow').ResultIU = 4
            sh.CellsU('BeginArrowSize').ResultIU = head_size
        return sh

    def line(self, pts, **kw):
        kw['head'] = 'none'
        return self.arrow(pts, **kw)

    # ---- 组合元素 ----
    def region(self, x0, y0, x1, y1, name, color=C.BLUE, size=18, gap=7, weight=2.5):
        """功能分区：彩色粗直角框，区名用同色加粗大字放在框下方。"""
        sh = self.box(x0, y0, x1, y1, line=color, weight=weight, round_mm=0)
        self.label((x0 + x1) / 2, y1 + gap, name, size=size, color=color, w=x1 - x0)
        return sh

    def group(self, x0, y0, x1, y1, color=C.RED, name=None, size=14, weight=2):
        """子模块组：彩色圆角虚线框，可在框上方标名称。"""
        sh = self.box(x0, y0, x1, y1, line=color, weight=weight, pattern='round_dash', round_mm=3)
        if name:
            self.label((x0 + x1) / 2, y0 - 4.5, name, size=size, color=color, w=x1 - x0)
        return sh

    def memory(self, x0, y0, x1, y1, txt, size=14):
        """存储器类模块：深灰填充+白色加粗字。"""
        return self.box(x0, y0, x1, y1, txt, fill=C.DARK, size=size, color=C.WHITE)

    def block_arrow(self, x, y, direction='down', length=8, width=6, color=C.RED):
        """红色粗块状箭头（常用于“识别结果”等最终输出），(x,y)为箭尾中点。"""
        d = {'down': (0, 1), 'up': (0, -1), 'right': (1, 0), 'left': (-1, 0)}[direction]
        ux, uy = d
        px, py = -uy, ux
        hl, sw, hw = length * 0.45, width * 0.45, width / 2
        tail = length - hl
        pts = [(x + px * sw, y + py * sw), (x + ux * tail + px * sw, y + uy * tail + py * sw),
               (x + ux * tail + px * hw, y + uy * tail + py * hw), (x + ux * length, y + uy * length),
               (x + ux * tail - px * hw, y + uy * tail - py * hw), (x + ux * tail - px * sw, y + uy * tail - py * sw),
               (x - px * sw, y - py * sw), (x + px * sw, y + py * sw)]
        xy = []
        for a, b in pts:
            xy += [self._x(a), self._y(b)]
        sh = self.page.DrawPolyline(xy, 0)
        sh.CellsU('FillPattern').ResultIU = 1
        sh.CellsU('FillForegnd').FormulaU = color
        sh.CellsU('LineColor').FormulaU = color
        sh.CellsU('LineWeight').FormulaU = '0.75 pt'
        return sh

    @staticmethod
    def char_style(sh, begin, end, style=BOLD_ITALIC, sub=False, sup=False):
        """设置文字片段格式（0起始，[begin,end)）：变量加粗斜体、下标/上标。"""
        ch = sh.Characters
        ch.Begin, ch.End = begin, end
        ch.CharProps(2, style)                 # visCharacterStyle
        if sub or sup:
            ch.CharProps(4, 2 if sub else 1)   # visCharacterPos：1上标 2下标
