package com.ms.createreport;

import com.ms.util.poi.FileDownload;
import com.ms.util.poi.MSWordTool;
import com.ms.util.poi.PropertiesUtil;
import net.sf.json.JSONArray;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.text.SimpleDateFormat;
import java.util.*;


/**
 * app检测司法鉴定意见书
 *
 * @author Masai
 * @date 2019-03-07
 */
@Controller
@RequestMapping(value = "/poi")
public class CreateAppReport {

    /** 生成报告主程序
     * @param request
     * @return
     */
    @RequestMapping(value = "createReport")
    public void getAppReport(HttpServletRequest request, HttpServletResponse response) throws Exception {
        String reportPath = PropertiesUtil.getValue("report.properties", "reportPath");
        String reportSavePath= PropertiesUtil.getValue("report.properties", "reportSavePath")+"测试审批表_"+File.separator;
        String reprotName = "司法鉴定意见书";
        MSWordTool changer = new MSWordTool();
        changer.setTemplateReturnDoc(reportPath);

        //获取当前时间
        SimpleDateFormat date = new SimpleDateFormat("yyyy年MM月dd日");
        Date now = new Date();
        String createTime = date.format(now);
        //获取传参
        String acceptData = request.getParameter("acceptData");
        String appName = request.getParameter("appName");
        String basicCase = request.getParameter("basicCase");
        String docNO = request.getParameter("docNO");
        String endTime = request.getParameter("endTime");
        String fileMD5 = request.getParameter("fileMD5");
        String md5 = request.getParameter("MD5");
        String fileName = request.getParameter("fileName");
        String fileSize = request.getParameter("fileSize");
        String matter = request.getParameter("matter");
        String sourceURL = request.getParameter("sourceURL");
        String startTime = request.getParameter("startTime");
        //获取图片
        String imgSet = request.getParameter("imgSet");
        JSONArray array = JSONArray.fromObject(imgSet);
        System.out.println(array.get(0));


        Map<String,String> map = new HashMap<>(16);
        map.put("备案编号", docNO);
        changer.replaceBookMarkText(map,false,false,10,"黑体");

        Map<String,String> map2 = new HashMap<>(16);
        map2.put("受理日期", acceptData);
        map2.put("鉴定事项", matter);
        map2.put("基本案情", basicCase);
        map2.put("鉴定时间", startTime+"至"+endTime);
        map2.put("鉴定过程时间", startTime+"至"+endTime);
        map2.put("哈希值", md5);
        map2.put("报告日期", createTime);
        changer.replaceBookMarkText(map2,false,false,14,"仿宋_GB2312");

        List<Map<String,String>> summaryMapList = new ArrayList<>();
        Map<String,String> map3 = new HashMap<>(16);
        map3.put("来源网址", sourceURL);
        map3.put("应用名称", appName);
        map3.put("安装包文件名", fileName);
        map3.put("大小", fileSize);
        map3.put("MD5", fileMD5);
        summaryMapList.add(map3);
        changer.fillTableAtBookMark("资料摘要", summaryMapList);

        map3.remove("来源网址", sourceURL);
        changer.fillTableAtBookMark("分析说明", summaryMapList);

        //到服务器存档,下载到默认浏览器下载的地方
        changer.saveAs(reportSavePath+appName+reprotName,reportSavePath);
        FileDownload.fileDownload(response,reportSavePath+appName+fileName, appName+"-"+fileName);

        //return "";
    }
}
