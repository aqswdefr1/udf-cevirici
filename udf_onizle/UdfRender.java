import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.*;
import javax.imageio.ImageIO;
import javax.swing.*;
import javax.swing.text.JTextComponent;

/** UDF'yi UYAP Doküman Editörü'nün kendi sınıflarıyla, pencere açmadan PNG'ye basar.
 *  Kullanım: java -cp ".:<editör jar'ları>" UdfRender girdi.udf cikti-öneki [ölçek] [tek]
 *  Çıktı: <önek>-sayfa-1.png, -2.png …  ("tek" verilirse ya da sayfa ölçülemezse <önek>-tum.png)
 *  Editör bir sayfanın zeminini/üstbilgisini yalnız o sayfa görünür alandayken çizer; bu yüzden
 *  sayfa yüksekliği ve aralığı ilk iki banttan ölçülür, her sayfa tam görünürken ayrı boyanır.   */
public class UdfRender {
    static JTextComponent best = null;
    static JFrame f;
    static JViewport vp;
    static int cw, vh;

    static void find(Container c) {
        for (Component ch : c.getComponents()) {
            if (ch instanceof JTextComponent) {
                JTextComponent t = (JTextComponent) ch;
                if (best == null || t.getDocument().getLength() > best.getDocument().getLength()) best = t;
            }
            if (ch instanceof Container) find((Container) ch);
        }
    }

    static boolean teal(int rgb) {
        int r = (rgb >> 16) & 255, g = (rgb >> 8) & 255, b = rgb & 255;
        return Math.abs(r - 44) < 16 && Math.abs(g - 153) < 16 && Math.abs(b - 174) < 16;
    }

    /** Görünüm alanını vy'ye kaydırıp o bandı (cw x vh) boyar; görüntünün y=0'ı belgenin vy'sidir. */
    static BufferedImage band(int vy) {
        vp.setViewPosition(new Point(0, vy));
        f.validate();
        BufferedImage im = new BufferedImage(cw, vh, BufferedImage.TYPE_INT_RGB);
        Graphics2D g = im.createGraphics();
        g.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
        g.setColor(new Color(44, 153, 174)); g.fillRect(0, 0, cw, vh);
        g.translate(0, -vy);
        g.setClip(0, vy, cw, vh);
        best.printAll(g);
        g.dispose();
        return im;
    }

    /** x sütununda, from'dan başlayarak ilk "sayfa" (turkuaz olmayan, >=100 px) aralığını bulur. */
    static int[] pageRun(BufferedImage im, int x, int from) {
        int y = from;
        while (y < im.getHeight() && teal(im.getRGB(x, y))) y++;
        int s = y;
        while (y < im.getHeight() && !teal(im.getRGB(x, y))) y++;
        return (y - s >= 100) ? new int[]{s, y} : null;
    }

    static void save(BufferedImage src, double scale, String path) throws IOException {
        BufferedImage img = src;
        if (scale != 1.0) {
            int w = (int) Math.round(src.getWidth() * scale), h = (int) Math.round(src.getHeight() * scale);
            img = new BufferedImage(w, h, BufferedImage.TYPE_INT_RGB);
            Graphics2D g = img.createGraphics();
            g.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BICUBIC);
            g.drawImage(src, 0, 0, w, h, null);
            g.dispose();
        }
        ImageIO.write(img, "png", new File(path));
    }

    public static void main(String[] a) throws Exception {
        final String in = a[0], out = a[1];
        final double scale = a.length > 2 ? Double.parseDouble(a[2]) : 1.0;
        final boolean single = a.length > 3 && a[3].equals("tek");
        final int[] rc = {1};
        SwingUtilities.invokeAndWait(new Runnable() { public void run() { try {
            tr.com.havelsan.uyap.system.editor.common.gui.WPDocumentPanel p =
                new tr.com.havelsan.uyap.system.editor.common.gui.WPDocumentPanel();
            boolean ok = p.a(new FileInputStream(in), in);
            f = new JFrame();
            f.getContentPane().add(p);
            f.setSize(1300, 1700);
            f.addNotify();
            f.validate();
            find(p);
            if (best == null) { System.err.println("UDFRENDER-HATA: metin bileşeni bulunamadı"); return; }
            Dimension d = best.getPreferredSize();
            cw = Math.max(d.width, best.getWidth());
            vp = (JViewport) SwingUtilities.getAncestorOfClass(JViewport.class, best);
            vh = vp == null ? d.height : Math.max(300, vp.getExtentSize().height);

            BufferedImage full = new BufferedImage(cw, d.height, BufferedImage.TYPE_INT_RGB);
            Graphics2D fg = full.createGraphics();
            fg.setColor(new Color(44, 153, 174)); fg.fillRect(0, 0, cw, d.height);
            int pages = 0;
            BufferedImage b0 = band(0);
            int px = -1;                                         // sayfanın içinden geçen bir sütun
            for (int x = 0; x < cw; x++) if (!teal(b0.getRGB(x, Math.min(60, vh - 1)))) { px = x + 4; break; }
            int[] r0 = px < 0 ? null : pageRun(b0, px, 0);
            if (vp == null || r0 == null || r0[1] >= vh - 2) {   // ölçülemedi: düz bantlama
                for (int y = 0; y < d.height; y += vh) {
                    int vy = Math.max(0, Math.min(y, d.height - vh));
                    fg.drawImage(band(vy), 0, vy, null);
                }
            } else {
                int top = r0[0], pageH = r0[1] - r0[0], stride = 0;
                if (r0[1] + 40 < d.height) {
                    int vy = Math.min(r0[1] + 2, Math.max(0, d.height - vh));
                    int[] r1 = pageRun(band(vy), px, 0);
                    if (r1 != null) stride = vy + r1[0] - top;
                }
                if (stride <= pageH) stride = pageH + 22;
                int xr = px;                                     // sayfanın sağ kenarı
                for (int x = cw - 1; x > px; x--) if (!teal(b0.getRGB(x, Math.min(60, vh - 1)))) { xr = x; break; }
                int x0 = Math.max(0, px - 8), x1 = Math.min(cw, xr + 5);
                for (int pt = top; pt + 50 < d.height; pt += stride) {
                    int vy = Math.max(0, Math.min(pt - 20, d.height - vh));
                    BufferedImage b = band(vy);
                    int y0 = Math.max(pt - 6, vy), y1 = Math.min(Math.min(pt + pageH + 6, vy + vh), d.height);
                    fg.drawImage(b.getSubimage(0, y0 - vy, cw, y1 - y0), 0, y0, null);
                    pages++;
                    if (!single) {
                        String pth = out + "-sayfa-" + pages + ".png";
                        save(b.getSubimage(x0, y0 - vy, x1 - x0, y1 - y0), scale, pth);
                        System.err.println("UDFRENDER-SAYFA " + pth);
                    }
                }
            }
            fg.dispose();
            if (single || pages == 0) {
                save(full, scale, out + "-tum.png");
                System.err.println("UDFRENDER-SAYFA " + out + "-tum.png");
            }
            System.err.println("UDFRENDER-OK yukleme=" + ok + " boyut=" + cw + "x" + d.height
                + " sayfa=" + pages + " uzunluk=" + best.getDocument().getLength());
            rc[0] = 0;
        } catch (Throwable t) { System.err.println("UDFRENDER-HATA: " + t); t.printStackTrace(); } } });
        System.exit(rc[0]);
    }
}
