package com.cesgroup.report.config.excel;

import java.util.HashMap;
import java.util.Map;

import com.cesgroup.report.config.excel.parse.DefaultParamParse;
import com.cesgroup.report.config.excel.parse.IParse;
import com.cesgroup.report.config.excel.parse.ListParamParse;

public class ParamParseMaker {
	
    private Map<Type, IParse> dict = new HashMap<Type, IParse>();

    private static ParamParseMaker factory = new ParamParseMaker();	
	
    private ParamParseMaker()
    {
    }

    public static ParamParseMaker getInstance()
    {
        return factory;
    }
    
    public IParse createParse(String expression)
    {
        IParse parse = null;
        
        Type type = Type.DEFAULT;
        
        if(expression.indexOf("param[") >= 0) {
        	type = Type.LIST;
        } 

        if (!dict.containsKey(type))
        {

            switch (type)
            {

                case LIST: parse = new ListParamParse(); break;
                default: parse = new DefaultParamParse(); break;


            }

            dict.put(type, parse);
        }

        return dict.get(type);
    }
    
	 enum Type
     {
         DEFAULT,
         HEADER,
         LIST
     }
}
