package ser;

import com.ser.blueline.IDocument;
import com.ser.blueline.IDocumentServer;
import com.ser.blueline.IInformationObject;
import com.ser.blueline.ISession;
import com.ser.blueline.bpm.IBpmService;
import com.ser.blueline.bpm.IProcessInstance;
import com.ser.blueline.bpm.ITask;
import com.spire.pdf.conversion.compression.ImageCompressionOptions;
import com.spire.xls.ImageFormatType;
import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;
import com.spire.xls.core.spreadsheet.HTMLOptions;
import de.ser.doxis4.agentserver.UnifiedAgent;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Calendar;
import java.util.Collection;
import java.util.Date;
import java.util.List;
import java.util.UUID;
import java.util.concurrent.TimeUnit;

import static java.lang.System.out;


public class TransmittalTest extends UnifiedAgent {
    ISession session;
    IDocumentServer server;
    IBpmService bpm;
    IProcessInstance processInstance;
    IInformationObject projectInfObj;
    ITask task;
    private ProcessHelper helper;
    @Override
    protected Object execute() {
        if (getEventDocument() == null)
            return resultError("Null Document object");

        session = getSes();
        server = session.getDocumentServer();
        IDocument document = getEventDocument();

        try {

            helper = new ProcessHelper(getSes());
            (new File(Conf.ExcelTransmittalPaths.MainPath)).mkdirs();

            XTRObjects.setSession(session);

            String uniqueId = UUID.randomUUID().toString();
            String excelPath = FileEvents.fileExport(document, Conf.ExcelTransmittalPaths.MainPath, uniqueId);

            FileInputStream fist = new FileInputStream(excelPath);
            XSSFWorkbook fwrb = new XSSFWorkbook(fist);

            JSONObject ecfg = Utils.getExcelConfig(fwrb);
            JSONObject data = Utils.getDataOfTransmittal(fwrb, ecfg);
            List<JSONObject> dist = Utils.getListOfDistributions(fwrb, ecfg);

        } catch (Exception e) {
            //throw new RuntimeException(e);
            out.println("Exception       : " + e.getMessage());
            out.println("    Class       : " + e.getClass());
            out.println("    Stack-Trace : " + e.getStackTrace() );
            return resultError("Exception : " + e.getMessage());
        }

        out.println("Finished");
        return resultSuccess("Ended successfully");
    }
    private static String
        loadStampExcel(String templateXlsxPath, String xslxPath, JSONObject bookmarks)throws Exception{

        FileInputStream tist = new FileInputStream(templateXlsxPath);
        XSSFWorkbook twrb = new XSSFWorkbook(tist);


        Sheet tsht = twrb.getSheetAt(0);
        for (Row trow : tsht){
            for(Cell tcll : trow){
                if(tcll.getCellType() != CellType.STRING){continue;}
                String clvl = tcll.getRichStringCellValue().getString();
                String clvv = Utils.updateCell(clvl, bookmarks);
                if(!clvv.equals(clvl)){
                    tcll.setCellValue(clvv);
                }
            }
        }
        FileOutputStream tost = new FileOutputStream(xslxPath);
        twrb.write(tost);
        tost.close();

        return xslxPath;
    }
    private static String
        stampImage(String xlsxPath, String pngPath)throws Exception{

        com.spire.xls.Workbook workbook = new Workbook();
        workbook.loadFromFile(xlsxPath);
        Worksheet sheet = workbook.getWorksheets().get(0);

        sheet.saveToImage(pngPath);
        return pngPath;
    }
    private static String
        transparency(String in, String out)throws Exception{
        BufferedImage bi = ImageIO.read(new File(in));
        int[] pixels = bi.getRGB(0, 0, bi.getWidth(), bi.getHeight(), null, 0, bi.getWidth());

        for(int i=0;i<pixels.length;i++){
            int color = pixels[i];
            int a = (color>>24)&255;
            int r = (color>>16)&255;
            int g = (color>>8)&255;
            int b = (color)&255;

            if(r == 255 && g == 255 && b == 255){
                a = 0;
            }

            pixels[i] = (a<<24) | (r<<16) | (g<<8) | (b);
        }

        BufferedImage biOut = new BufferedImage(bi.getWidth(), bi.getHeight(), BufferedImage.TYPE_INT_ARGB);
        biOut.setRGB(0, 0, bi.getWidth(), bi.getHeight(), pixels, 0, bi.getWidth());
        ImageIO.write(biOut, "png", new File(out));
        return out;
    }
    public static String
        autoCrop(String in, String out, double tolerance) throws Exception{
        BufferedImage source = ImageIO.read(new File(in));

        int baseColor = source.getRGB(0, 0);

        int width = source.getWidth();
        int height = source.getHeight();

        int minX = 0;
        int minY = 0;
        int maxX = width;
        int maxY = height;
        int margin = 10;

        // Immediately break the loops when encountering a non-white pixel.
        lable1: for (int y = 0; y < height; y++) {
            for (int x = 0; x < width; x++) {
                if (colorWithinTolerance(baseColor, source.getRGB(x, y), tolerance)) {
                    minY = y;
                    break lable1;
                }
            }
        }

        lable2: for (int x = 0; x < width; x++) {
            for (int y = minY; y < height; y++) {
                if (colorWithinTolerance(baseColor, source.getRGB(x, y), tolerance)) {
                    minX = x;
                    break lable2;
                }
            }
        }

        baseColor = source.getRGB(minX, height - 1);

        lable3: for (int y = height - 1; y >= minY; y--) {
            for (int x = minX; x < width; x++) {
                if (colorWithinTolerance(baseColor, source.getRGB(x, y), tolerance)) {
                    maxY = y;
                    break lable3;
                }
            }
        }

        lable4: for (int x = width - 1; x >= minX; x--) {
            for (int y = minY; y < maxY; y++) {
                if (colorWithinTolerance(baseColor, source.getRGB(x, y), tolerance)) {
                    maxX = x;
                    break lable4;
                }
            }
        }

        if ((minX - margin) >= 0) {
            minX -= margin;
        }

        if ((minY - margin) >= 0) {
            minY -= margin;
        }

        if ((maxX + margin) < width) {
            maxX += margin;
        }

        if ((maxY + margin) < height) {
            maxY += margin;
        }

        int newWidth = maxX - minX + 1;
        int newHeight = maxY - minY + 1;

        if (newWidth == width && newHeight == height) {
            ImageIO.write(source, "png", new File(out));
            return out;
        }

        BufferedImage target = new BufferedImage(newWidth, newHeight, source.getType());

        Graphics g = target.getGraphics();
        ((Graphics) g).drawImage(source, 0, 0, target.getWidth(), target.getHeight(), minX, minY, maxX + 1, maxY + 1, null);

        g.dispose();

        ImageIO.write(target, "png", new File(out));
        return out;
    }
    private static boolean colorWithinTolerance(int a, int b, double tolerance) {
        int aAlpha = (int) ((a & 0xFF000000) >>> 24); // Alpha level
        int aRed = (int) ((a & 0x00FF0000) >>> 16); // Red level
        int aGreen = (int) ((a & 0x0000FF00) >>> 8); // Green level
        int aBlue = (int) (a & 0x000000FF); // Blue level

        int bAlpha = (int) ((b & 0xFF000000) >>> 24); // Alpha level
        int bRed = (int) ((b & 0x00FF0000) >>> 16); // Red level
        int bGreen = (int) ((b & 0x0000FF00) >>> 8); // Green level
        int bBlue = (int) (b & 0x000000FF); // Blue level

        double distance = Math.sqrt((aAlpha - bAlpha) * (aAlpha - bAlpha) + (aRed - bRed) * (aRed - bRed)
                + (aGreen - bGreen) * (aGreen - bGreen) + (aBlue - bBlue) * (aBlue - bBlue));

        double percentAway = distance / 510.0d;

        return (percentAway > tolerance);
    }
}