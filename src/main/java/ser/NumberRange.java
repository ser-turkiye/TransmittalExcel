package ser;

import com.ser.blueline.IDocumentServer;
import com.ser.blueline.ISession;
import com.ser.blueline.metaDataComponents.IStringMatrix;
import com.ser.blueline.modifiablemetadata.IStringMatrixModifiable;
import com.spire.ms.System.Exception;
import org.json.JSONObject;
import ser.XTRObjects;

import java.math.BigInteger;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

public class NumberRange {
    ISession session;
    IDocumentServer server;
    String principalName;
    IStringMatrix matrix;
    JSONObject numberRanges;
    AutoText autoText;

    int columnName = 0;
    int columnPattern = 1;
    int columnStart = 2;
    int columnId = 3;
    int columnCurrent = 4;
    public NumberRange() throws Exception {
        load();
    }
    public void load() throws Exception {
        session = XTRObjects.getSession();
        if(session == null){throw new Exception("Please set session with XTRObjects.init function.");}

        server = XTRObjects.getServer();
        principalName = XTRObjects.getPrincipalName();
        matrix = server.getStringMatrix("SYS_NUMBER_RANGES", session);

        if(matrix == null){throw new Exception("SYS_NUMBER_RANGES Global Value List not found");}

        numberRanges = new JSONObject();
        autoText = new AutoText();
        List<List<String>> rmtx = matrix.getRawRows();

        int lcnt = (-1);
        for(List<String> line : rmtx) {
            lcnt++;
            String ixnm = (line.size() > columnName && line.get(columnName) != null ? line.get(columnName) : "");
            String ixpr = (line.size() > columnPattern && line.get(columnPattern) != null ? line.get(columnPattern) : "");

            if (ixnm.isEmpty() || ixpr.isEmpty()) {continue;}
            if (numberRanges.has(ixnm)) {continue;}

            String ixst = (line.size() > columnStart && line.get(columnStart) != null ? line.get(columnStart) : "");
            long lxst = (ixst == "" ? 0L : Long.parseLong(ixst));
            lxst = (lxst < 0L ? 0L : lxst);

            String ixid = (line.size() > columnId && line.get(columnId) != null ? line.get(columnId) : "");
            long lxid = (ixid == "" ? 0L : Long.parseLong(ixid));

            JSONObject nrng = new JSONObject();
            nrng.put("$", lcnt);
            nrng.put("name", ixnm);
            nrng.put("pattern", ixpr);
            nrng.put("start", lxst);
            nrng.put("id", lxid);

            numberRanges.put(ixnm, nrng);
        }
        for (String nrky : numberRanges.keySet()) {
            JSONObject nrng = (JSONObject) numberRanges.get(nrky);
            if(!nrng.has("$") || nrng.getInt("$") < 0 || rmtx.get(nrng.getInt("$")) == null){continue;}

            int rown = nrng.getInt("$");
            if(nrng.getLong("id") != 0L){
                counter(nrky, false);
                continue;
            }

            long lnid = longId(principalName + "\\" + nrky);

            IStringMatrixModifiable mmtx = matrix.getModifiableCopy(session);
            mmtx.setValue(rown, columnId, Long.toString(lnid), false);
            mmtx.commit();
            matrix.refresh();

            session.deleteCounter(lnid);
            session.getNextCounterValue(lnid, nrky, nrng.getLong("start"), 1L);

            nrng.put("id", lnid);
            numberRanges.put(nrky, nrng);
        }
    }
    public static NumberRange init() throws Exception {
        return new NumberRange();
    }
    public NumberRange with(Object obj) throws Exception {
        autoText.with(obj);
        return this;
    }
    public void parameter(String pnam, String pval) throws java.lang.Exception {
        autoText.parameter(pnam, pval);
    }
    public NumberRange param(String pnam, String pval) throws java.lang.Exception {
        autoText.parameter(pnam, pval);
        return this;
    }
    public String increment(String name) throws Exception {
        try {
            return format(name, incrementLong(name));
        } catch (java.lang.Exception e) {
            throw new Exception(e.toString());
        }
    }
    public void append(String name, String pattern, long start) throws Exception {
        if(has(name)){throw new Exception("Number range '" + name + "' is exists.");}

        int rwnr = matrix.getRowCount();
        IStringMatrixModifiable mmtx = matrix.getModifiableCopy(session);
        mmtx.appendRow();
        mmtx.setValue(rwnr, columnName, name, false);
        mmtx.setValue(rwnr, columnPattern, pattern, false);
        mmtx.setValue(rwnr, columnStart, Long.toString(Math.max(start, 0L)), false);
        mmtx.commit();
        matrix.refresh();

        load();
    }
    public String current(String name) throws Exception {
        try {
            return format(name, currentLong(name));
        } catch (java.lang.Exception e) {
            throw new Exception(e.toString());
        }
    }
    public String format(String name, long count) throws java.lang.Exception {
        String rtrn = pattern(name);
        rtrn = (rtrn == "" ? "%n%" : rtrn);
        String cnt = ((Long) count).toString();
        int maxr = (cnt.length() < 16 ? 16 : cnt.length());

        for(int i=maxr;i>0;i--){
            if(cnt.length() > i){break;}

            String akey = "%n" + i + "%";
            if(!rtrn.contains(akey)){continue;}

            String aval = ("0".repeat(i) + cnt).substring(cnt.length());
            rtrn = rtrn.replaceAll(akey, aval);
        }
        Date date = new Date();
        for(String dfrl : XTRObjects.patternsDateTime){
            String dkey = "%" + dfrl + "%";
            if(!rtrn.contains(dkey)){continue;}
            String dval = (new SimpleDateFormat(dfrl)).format(date);
            rtrn = rtrn.replaceAll(dkey, dval);
        }

        rtrn = autoText.run(rtrn);
        return rtrn;
    }
    public boolean has(String id) throws Exception {
        return numberRanges.has(id);
    }
    public long incrementLong(String name) throws Exception {
        return counter(name, true);
    }
    public long currentLong(String name) throws Exception {
        return counter(name, false);
    }
    public String pattern(String name) throws Exception {
        if(!has(name)){throw new Exception("Number range '" + name + "' not found.");}
        JSONObject nrng = (JSONObject) numberRanges.get(name);
        return (nrng.has("pattern") ? nrng.getString("pattern") : "");
    }
    public long counter(String name, boolean incr) throws Exception {
        if(!has(name)){throw new Exception("Number range '" + name + "' not found.");}
        JSONObject nrng = (JSONObject) numberRanges.get(name);
        if(!nrng.has("id") || nrng.getLong("id") == 0L){throw new Exception("Number range '" + name + "' id info not found.");}

        long strt = nrng.getLong("start");
        long lnid = nrng.getLong("id");
        long rtrn = 0L;
        rtrn = session.getNextCounterValue(lnid, name, strt, 1L);
        if(!incr && rtrn > strt) {
            session.deleteCounter(lnid);
            rtrn = session.getNextCounterValue(lnid, name,  rtrn-1, 1L);
        }

        int rown = nrng.getInt("$");
        IStringMatrixModifiable mmtx = matrix.getModifiableCopy(session);
        mmtx.setValue(rown, columnCurrent, Long.toString(rtrn), false);
        mmtx.commit();
        matrix.refresh();

        nrng.put("current", rtrn);
        numberRanges.put(name, nrng);

        return rtrn;
    }
    public long longId(String text) throws Exception{
        byte[] md5hex = new byte[0];
        try {
            md5hex = MessageDigest.getInstance("MD5").digest(text.getBytes());
        } catch (NoSuchAlgorithmException e) {
            throw new Exception(e.toString());
        }
        return (new BigInteger(bytesToHex(md5hex), 16)).longValue();
    }
    public String bytesToHex(byte[] bytes) {
        char[] hexArray = "0123456789ABCDEF".toCharArray();
        char[] hexChars = new char[bytes.length * 2];
        for (int j = 0; j < bytes.length; j++) {
            int v = bytes[j] & 0xFF;
            hexChars[j * 2] = hexArray[v >>> 4];
            hexChars[j * 2 + 1] = hexArray[v & 0x0F];
        }
        return new String(hexChars);
    }
}
