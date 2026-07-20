// 生成 cmb_pdf2excel 应用图标（多尺寸 PNG 编码 ICO：256/48/32/16）
// 设计：现代渐变圆角背景 + 白色循环转换箭头（寓意 PDF → Excel 转换）
// 编译: csc /target:exe /out:scripts\make_icon.exe /r:System.Drawing.dll scripts\make_icon.cs
// 用法: make_icon.exe [输出.ico路径]   （默认 app.ico，同时生成同名 .preview.png 预览）
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;

class MakeIcon
{
    static void Main(string[] args)
    {
        string outPath = args.Length > 0 ? args[0] : "app.ico";
        int[] sizes = { 256, 48, 32, 16 };

        var pngs = new List<byte[]>();
        foreach (int s in sizes)
        {
            using (Bitmap bmp = Render(s))
            using (var ms = new MemoryStream())
            {
                bmp.Save(ms, ImageFormat.Png);
                pngs.Add(ms.ToArray());
            }
        }

        // 预览图（256）
        File.WriteAllBytes(Path.ChangeExtension(outPath, ".preview.png"), pngs[0]);

        using (var fs = File.Create(outPath))
        using (var w = new BinaryWriter(fs))
        {
            w.Write((short)0);              // reserved
            w.Write((short)1);              // type = icon
            w.Write((short)sizes.Length);   // image count
            int offset = 6 + 16 * sizes.Length;
            for (int i = 0; i < sizes.Length; i++)
            {
                int s = sizes[i];
                w.Write((byte)(s >= 256 ? 0 : s));  // width (0 => 256)
                w.Write((byte)(s >= 256 ? 0 : s));  // height
                w.Write((byte)0);           // color palette
                w.Write((byte)0);           // reserved
                w.Write((short)1);          // color planes
                w.Write((short)32);         // bits per pixel
                w.Write((int)pngs[i].Length);
                w.Write((int)offset);
                offset += pngs[i].Length;
            }
            foreach (var p in pngs) w.Write(p);
        }
        Console.WriteLine("Icon created: " + outPath);
    }

    // 按目标尺寸等比绘制
    static Bitmap Render(int size)
    {
        var bmp = new Bitmap(size, size);
        using (Graphics g = Graphics.FromImage(bmp))
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
            g.Clear(Color.Transparent);

            float sc = size / 256f;
            int margin = (int)(8 * sc);
            int radius = (int)(58 * sc);
            var rect = new Rectangle(margin, margin, size - 2 * margin, size - 2 * margin);

            // 渐变圆角背景（靛蓝 → 青，科技感 / 转换）
            using (var brush = new LinearGradientBrush(
                       rect, Color.FromArgb(79, 70, 229), Color.FromArgb(6, 182, 212),
                       LinearGradientMode.ForwardDiagonal))
            using (var path = RoundedRect(rect, radius))
            {
                g.FillPath(brush, path);
            }

            // 循环转换箭头
            float cx = size / 2f, cy = size / 2f;
            float R = 60 * sc;
            float penW = 21 * sc;
            using (var pen = new Pen(Color.White, penW))
            {
                pen.StartCap = LineCap.Round;
                pen.EndCap = LineCap.Round;
                g.DrawArc(pen, cx - R, cy - R, 2 * R, 2 * R, 220, 120);  // 上弧
                g.DrawArc(pen, cx - R, cy - R, 2 * R, 2 * R, 40, 120);   // 下弧
            }
            DrawArrowHead(g, cx, cy, R, 340, penW);  // 上弧箭头
            DrawArrowHead(g, cx, cy, R, 160, penW);  // 下弧箭头
        }
        return bmp;
    }

    // 在弧线末端（顺时针方向）绘制白色三角箭头
    static void DrawArrowHead(Graphics g, float cx, float cy, float R, float endDeg, float penW)
    {
        double a = endDeg * Math.PI / 180.0;
        float ex = (float)(cx + R * Math.Cos(a));
        float ey = (float)(cy + R * Math.Sin(a));
        double tang = a + Math.PI / 2;          // 顺时针切线方向
        float tx = (float)Math.Cos(tang);
        float ty = (float)Math.Sin(tang);
        float px = -ty, py = tx;                // 垂直方向

        float tipLen = penW * 1.75f;
        float halfW = penW * 1.3f;
        var tip = new PointF(ex + tx * tipLen, ey + ty * tipLen);
        var b1 = new PointF(ex + px * halfW, ey + py * halfW);
        var b2 = new PointF(ex - px * halfW, ey - py * halfW);

        using (var brush = new SolidBrush(Color.White))
            g.FillPolygon(brush, new[] { tip, b1, b2 });
    }

    static GraphicsPath RoundedRect(Rectangle b, int r)
    {
        int d = r * 2;
        var p = new GraphicsPath();
        p.AddArc(b.X, b.Y, d, d, 180, 90);
        p.AddArc(b.Right - d, b.Y, d, d, 270, 90);
        p.AddArc(b.Right - d, b.Bottom - d, d, d, 0, 90);
        p.AddArc(b.X, b.Bottom - d, d, d, 90, 90);
        p.CloseFigure();
        return p;
    }
}
