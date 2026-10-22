# 示例：JESD204B系统典型组成框图（时钟芯片 + ADC/FPGA/DAC三分区）
# 运行：python jesd204b_system.py [输出路径前缀]，默认输出到本文件同目录
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'scripts'))
from vsd_style import Figure, C, BOLD_ITALIC

out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), 'jesd204b_system')

with Figure(352, 128, out) as f:
    # 时钟芯片（关键模块：浅粉）
    f.box(140, 18, 210, 40, fill=C.PINK, weight=2.5)
    f.label(175, 25, '时钟芯片', 18, w=60)
    f.label(175, 34, '（HMC7044/LMK04828）', 11, w=68)
    f.label(175, 5, '参考时钟', w=30)
    f.arrow([(175, 9.5), (175, 18)])

    # 功能分区：彩色粗框，区名在框下
    f.region(28, 76, 98, 114, 'ADC', C.BLUE)
    f.region(125, 76, 225, 114, 'FPGA', C.RED)
    f.region(252, 76, 322, 114, 'DAC', C.BLUE)

    # 内部模块：灰阶填充
    f.box(32, 82, 54, 108, 'ADC\n内核', fill=C.GRAY1)
    f.box(64, 82, 94, 108, 'JESD204B\n发送端', fill=C.GRAY2)
    f.box(129, 82, 159, 108, 'JESD204B\n接收IP', fill=C.GRAY2)
    f.box(165, 82, 185, 108, '用户\n逻辑', fill=C.LGREEN)
    f.box(191, 82, 221, 108, 'JESD204B\n发送IP', fill=C.GRAY2)
    f.box(256, 82, 286, 108, 'JESD204B\n接收端', fill=C.GRAY2)
    f.box(296, 82, 318, 108, 'DAC\n内核', fill=C.GRAY1)
    for a, b in [(54, 64), (159, 165), (185, 191), (286, 296)]:
        f.arrow([(a, 95), (b, 95)])

    # 外部输入/输出从左右两侧进出
    f.arrow([(4, 95), (32, 95)]); f.label(15, 90, '模拟输入', w=26)
    f.arrow([(318, 95), (348, 95)]); f.label(336, 90, '模拟输出', w=26)

    # 高速lane（粗线）与SYNC~（绿色虚线）
    for a, b in [(94, 129), (221, 256)]:
        c = (a + b) / 2
        f.arrow([(a, 88.5), (b, 88.5)], weight=4.5, head_size=0)
        f.char_style(f.label(c, 84, 'lane×L', 13, w=24), 5, 6, BOLD_ITALIC)
        f.arrow([(b, 101.5), (a, 101.5)], color=C.GREEN, pattern='dash')
        f.label(c, 106.5, 'SYNC~', 13, color=C.GREEN, w=24)

    # 器件时钟（黑色实线）与SYSREF（紫色虚线）
    for clk, ref in [([(140, 25), (72, 25), (72, 76)], [(140, 34), (82, 34), (82, 76)]),
                     ([(168, 40), (168, 76)], [(182, 40), (182, 76)]),
                     ([(210, 25), (278, 25), (278, 76)], [(210, 34), (268, 34), (268, 76)])]:
        f.arrow(clk)
        f.arrow(ref, color=C.PURPLE, pattern='dash')
        xc, xr = clk[-1][0], ref[-1][0]
        if xc < xr:   # 时钟线在左：标签分居两侧
            f.label(xc - 2, 58, '器件时钟', align='right', w=30)
            f.label(xr + 2, 58, 'SYSREF', color=C.PURPLE, align='left', w=30)
        else:
            f.label(xr - 2, 58, 'SYSREF', color=C.PURPLE, align='right', w=30)
            f.label(xc + 2, 58, '器件时钟', align='left', w=30)
print('saved', out)
