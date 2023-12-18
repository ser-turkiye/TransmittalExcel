package ser;

import com.ser.blueline.IDescriptor;
import com.ser.blueline.IInformationObject;
import org.json.JSONObject;
import ser.XTRObjects;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class AutoText {
    JSONObject params;
    List<Object> objs;
    public AutoText(){
        params = new JSONObject();
        objs = new ArrayList<Object>();
    }
    public static AutoText init() throws Exception {
        return new AutoText();
    }
    public AutoText with(Object object){
        if(object == null){return this;}
        List<String> dxcs =  new ArrayList<>(Arrays.asList(
                XTRObjects.classNameIInformationObject,
                XTRObjects.classNameIDocument,
                XTRObjects.classNameITask,
                XTRObjects.classNameIProcessInstance
        ));

        if(dxcs.contains(object.getClass().getName())) {
            objs.add(object);
        }
        return this;
    }
    public void parameter(String pnam, String pval) throws Exception{
        params.put(pnam, pval);
    }
    public AutoText param(String pnam, String pval) throws Exception{
        parameter(pnam, pval);
        return this;
    }
    public String  value(String pnam, String prp) throws Exception{
        IDescriptor d = XTRObjects.getDescriptor(pnam);
        if(d == null){return "";}

        int dtyp = d.getDescriptorType();

        for(Object o : objs){
            IInformationObject info = (IInformationObject) o;
            if(!XTRObjects.infoObject(info).hasDescriptor(pnam)){continue;}

            if(dtyp == IDescriptor.TYPE_STRING) {
                String pvalString = info.getDescriptorValue(pnam, String.class);
                return (pvalString == null ? "" : pvalString);
            }
            if(dtyp == IDescriptor.TYPE_DOUBLE) {
                Double pvalDouble = info.getDescriptorValue(pnam, Double.class);
                return (pvalDouble == null ? "" : Double.toString(pvalDouble));
            }
            if(dtyp == IDescriptor.TYPE_INTEGER) {
                Integer pvalInteger = info.getDescriptorValue(pnam, Integer.class);
                return (pvalInteger == null ? "" : Integer.toString(pvalInteger));
            }
            if(dtyp == IDescriptor.TYPE_LONG) {
                Long pvalLong = info.getDescriptorValue(pnam, Long.class);
                return (pvalLong == null ? "" : Long.toString(pvalLong));
            }
            if(dtyp == IDescriptor.TYPE_BOOLEAN) {
                Boolean pvalBoolean = info.getDescriptorValue(pnam, Boolean.class);
                return (pvalBoolean == null ? "" : Boolean.toString(pvalBoolean));
            }
            if(dtyp == IDescriptor.TYPE_DATETIME || dtyp == IDescriptor.TYPE_DATE) {
                Date pvalDate = info.getDescriptorValue(pnam, Date.class);
                if(pvalDate != null) {
                    if(prp.isEmpty() && !XTRObjects.patternDefaultDateTime.isEmpty()) {
                        String dtvlDefault = (new SimpleDateFormat(XTRObjects.patternDefaultDateTime)).format(pvalDate);
                        return (dtvlDefault == null ? "" : dtvlDefault);
                    }
                    if(!prp.isEmpty() && XTRObjects.patternsDateTime.contains(prp)){
                        String dtvlString = (new SimpleDateFormat(prp)).format(pvalDate);
                        return (dtvlString == null ? "" : dtvlString);
                    }
                }
            }

        }
        return "";
    }
    public String run(String str) throws Exception{
        StringBuffer rtr1 = new StringBuffer();
        String tmp = str + "";
        String spr = "/";
        Pattern ptr1 = Pattern.compile( "\\{([\\w\\.\\:\\,\\/\\[\\]]+)\\}" );
        Matcher mtc1 = ptr1.matcher(tmp);
        while(mtc1.find()) {
            String mk = mtc1.group(1);
            String mv = "";
            if(params.has(mk)){
                mv = params.getString(mk);
            }
            if(mv == "") {
                String mf = (mk.contains(spr) ? mk.substring(0, mk.indexOf(spr)) : mk);
                String mp = (mk.contains(spr) ? mk.substring(mk.indexOf(spr) + spr.length()) : "");
                mv = value(mf, mp);
            }
            mtc1.appendReplacement(rtr1,  mv);
        }
        mtc1.appendTail(rtr1);
        tmp = rtr1.toString();

        return tmp;
    }
}
