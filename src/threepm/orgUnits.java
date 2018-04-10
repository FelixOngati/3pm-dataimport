package threepm;

import java.util.HashMap;
import java.util.Map;
import java.util.Iterator;
import java.util.Set;

public class orgUnits {
    public  HashMap<Integer,String> organizationalUnit() {
        HashMap<Integer,String> hm=new HashMap<Integer, String>();
        hm.put(18291,"F1etBBfosCl");
        hm.put(265,"y1x13IydCG7");
        hm.put(315,"GjGpfHjSGbO");
        return hm;
    }
    public String getOrgUid(Integer mflCode){
        //TODO Make an API Call orgunits endpoint. Returns response that comes with uid
        String var= organizationalUnit().get(mflCode);
        return var;
    }


}
