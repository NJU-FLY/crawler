package excel;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.*;
import jxl.write.Number;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

/**
 * Created by Administrator on 2016/9/19.
 */
public class TableForReport {

    /**
     * 写入立项时间数据，这个代码基本是不能用的
     *
     * @throws Exception
     */
    public void addLiXiang() throws Exception {
        //读
        InputStream instream = new FileInputStream("resources/2015年书目.xls");
        Workbook readwb = Workbook.getWorkbook(instream);
        Sheet sheet = readwb.getSheet(1);
        Integer[] results = new Integer[(sheet.getRows() - 2)];
        String[] title = new String[(sheet.getRows() - 2)];
        String[] author = new String[(sheet.getRows() - 2)];
        for (int i = 2; i < sheet.getRows(); i++) {
            Cell liXiang = sheet.getCell(1, i);
            author[i - 2] = sheet.getCell(5, i).getContents();
            title[i - 2] = sheet.getCell(3, i).getContents();
            String time = liXiang.getContents();
            if (time.length() >= 3) {
                time = time.substring(0, 2);
                time = "20" + time;
                Integer number = Integer.parseInt(time);
                results[i - 2] = number;
                System.out.println(results[i - 2]);
            }
        }

        //写
        Workbook rwb = Workbook.getWorkbook(new File("resources/result.xls"));
        WritableWorkbook wwb = Workbook.createWorkbook(new File("resources/result.xls"), rwb);//copy
        WritableSheet ws = wwb.getSheet(0);
        Sheet resultSheet = rwb.getSheet(0);
        HashMap<String, String> resultAuthor = new HashMap<>();
        System.out.println(resultSheet.getRows());
        for (int i = 1; i < resultSheet.getRows(); i++) {
            String result16 = resultSheet.getCell(0, i).getContents();
            if (resultAuthor.get(result16) != null) {
                System.out.println(result16);
                System.out.println(resultAuthor.get(result16));
                System.out.println(i);
            }
            resultAuthor.put(result16, resultSheet.getCell(1, i).getContents());
        }
        int rows = 2;
        for (int i = 1; i <= results.length; i++) {
            for (Iterator<Map.Entry<String, String>> iterator = resultAuthor.entrySet().iterator(); iterator.hasNext(); ) {
                Map.Entry<String, String> it = iterator.next();
                if (author[i - 1].contains(it.getValue()) && title[i - 1].contains(it.getKey())) {
                    Number num = new Number(23, rows, results[i - 1]);
                    ws.addCell(num);
                    break;
                } else if (author[i - 1].contains(it.getValue())) {
                    System.out.println(title[i - 1] + " " + author[i - 1] + " " + results[i - 1]);
                }
            }
            rows++;
        }
        wwb.write();
        wwb.close();
        rwb.close();
    }

    /**
     * 这个年均引用是用来生成年均引用数据列的，因为图书的年均引用长期来看是肯定下降的，所以已经不再需要这个数据了
     *
     * @param path
     * @throws IOException
     * @throws BiffException
     * @throws WriteException
     */
    public void 计算年均引用(String path) throws IOException, BiffException, WriteException {
        Sheet sheet = this.readSheet(path, 0);
        Integer rows = sheet.getRows();
        Double[] results = new Double[rows - 1];
        for (int i = 1; i < rows; i++) {
            Integer duration = null;
            Integer year = 2007;
            Double total = 0.0;
            Cell[] cells = new Cell[CellName.cit2007.getValue() - CellName.cit2015.getValue() + 1];
            for (int j = CellName.cit2007.getValue(); j > CellName.cit2015.getValue(); j--) {
                cells[j - CellName.cit2015.getValue()] = sheet.getCell(j, i);
                String cellStr = cells[j - CellName.cit2015.getValue()].getContents();
                if (!cellStr.equals("") && duration == null) {
                    System.out.println(year);
                    duration = 2015 - year;

                } else if (!cellStr.equals("")) {
                    total += Integer.parseInt(cellStr);
                }
                year++;
            }
            if (duration == null) {
                results[i - 1] = null;
            } else {
                results[i - 1] = total / duration;
            }
            System.out.println(results[i - 1]);
        }

        WritableWorkbook wwb = this.writeSheet(path);
        WritableSheet ws = wwb.getSheet(0);
        for (int i = 1; i < ws.getRows(); i++) {
            if (results[i - 1] != null) {
                Number cell = new Number(CellName.年均引用.getValue(), i, results[i - 1]);
                ws.addCell(cell);
            } else {
                Label cell = new Label(CellName.年均引用.getValue(), i, "null");
                ws.addCell(cell);
            }
        }
        this.closeSheet(wwb);
    }

    /**
     * 计算首次被引时间列，2016最终数据表的AO列
     *
     * @param path
     * @throws IOException
     * @throws WriteException
     * @throws BiffException
     */
    public void 计算首次被引时间(String path) throws IOException, WriteException, BiffException {
        Sheet sheet = this.readSheet(path, 0);
        Integer rows = sheet.getRows();
        Integer[] results = new Integer[rows - 1];
        for (int i = 1; i < rows; i++) {
            Integer firstCitYear = null;
            Integer year = 2007;
            Cell[] cells = new Cell[CellName.cit2007.getValue() - CellName.cit2015.getValue() + 1];
            for (int j = CellName.cit2007.getValue(); j >= CellName.cit2015.getValue(); j--) {
                cells[j - CellName.cit2015.getValue()] = sheet.getCell(j, i);
                String cellStr = cells[j - CellName.cit2015.getValue()].getContents();
                if (!cellStr.equals("")) {
                    firstCitYear = year;
                    break;
                }
                year++;
            }

            if (firstCitYear == null || firstCitYear == 2016) {
                results[i - 1] = null;
            } else {
                results[i - 1] = year;
            }
            System.out.println(results[i - 1]);
        }

        WritableWorkbook wwb = this.writeSheet(path);
        WritableSheet ws = wwb.getSheet(0);
        for (int i = 1; i < ws.getRows(); i++) {
            if (results[i - 1] != null) {
                Number cell = new Number(CellName.首次被引时间.getValue(), i, results[i - 1]);
                ws.addCell(cell);
            } else {
                Label cell = new Label(CellName.首次被引时间.getValue(), i, "null");
                ws.addCell(cell);
            }
        }
        this.closeSheet(wwb);
    }

    /**
     * 计算峰值间隔列，2016最终数据表的AN列
     *
     * @param path
     * @throws IOException
     * @throws BiffException
     * @throws WriteException
     */
    public void 计算峰值间隔(String path) throws IOException, BiffException, WriteException {
        Sheet sheet = this.readSheet(path, 0);
        Integer rows = sheet.getRows();
        Integer[] results = new Integer[rows - 1];
        for (int i = 1; i < rows; i++) {
            Integer max = 0;
            Integer maxYear = null;
            Integer year = 2015;
            Integer pubYear = Integer.parseInt(sheet.getCell(3, i).getContents());
            Cell[] cells = new Cell[CellName.cit2007.getValue() - CellName.cit2015.getValue() + 1];
            for (int j = CellName.cit2015.getValue(); j <= CellName.cit2007.getValue(); j++) {
                cells[j - CellName.cit2015.getValue()] = sheet.getCell(j, i);
                String cellStr = cells[j - CellName.cit2015.getValue()].getContents();
                Integer cellVal;
                if (cellStr.equals("")) {
                    cellVal = 0;
                } else {
                    cellVal = Integer.parseInt(cellStr);
                }
                if (cellVal >= max) {
                    max = cellVal;
                    maxYear = year;
                }
                year--;
            }
            if (max == 0) {
                maxYear = null;
            }

            if (maxYear == null) {
                results[i - 1] = null;
            } else {
                results[i - 1] = maxYear - pubYear;
            }
            System.out.println(results[i - 1]);
        }

        WritableWorkbook wwb = this.writeSheet(path);
        WritableSheet ws = wwb.getSheet(0);
        for (int i = 1; i < ws.getRows(); i++) {
            if (results[i - 1] != null) {
                Number cell = new Number(CellName.峰值间隔.getValue(), i, results[i - 1]);
                ws.addCell(cell);
            } else {
                Label cell = new Label(CellName.峰值间隔.getValue(), i, "null");
                ws.addCell(cell);
            }
        }
        this.closeSheet(wwb);
    }

    /**
     * 统计开头的函数，是用于生成统计表格的，这个是2016报告里的表9
     * 需要注意的是立项时间的截止时间是写死的，需要处理
     *
     * @param inputPath
     * @param outputPath
     * @throws Exception
     */
    public void 统计各种评论数(String inputPath, String outputPath) throws Exception {
        Sheet sheet = this.readSheet(inputPath, 0);
        Integer rows = sheet.getRows();
        String[] projectTime = new String[rows - 1];
        String[] doubanCount = new String[rows - 1];
        String[] doubanScore = new String[rows - 1];
        String[] dangdangCount = new String[rows - 1];
        String[] amazonCount = new String[rows - 1];
        String[] amazonScore = new String[rows - 1];
        String[] scholarCount = new String[rows - 1];

        for (int i = 1; i < rows; i++) {
            projectTime[i - 1] = sheet.getCell(CellName.projectTime.getValue(), i).getContents();
            doubanCount[i - 1] = sheet.getCell(CellName.doubanCommentCount.getValue(), i).getContents();
            doubanScore[i - 1] = sheet.getCell(CellName.doubanScore.getValue(), i).getContents();
            dangdangCount[i - 1] = sheet.getCell(CellName.dangdangCommentCount.getValue(), i).getContents();
            amazonCount[i - 1] = sheet.getCell(CellName.amazonCommentCount.getValue(), i).getContents();
            amazonScore[i - 1] = sheet.getCell(CellName.amazonScore.getValue(), i).getContents();
            scholarCount[i - 1] = sheet.getCell(CellName.scholarCommentCount.getValue(), i).getContents();
        }
        WritableWorkbook wwb = this.writeSheet(outputPath);
        WritableSheet ws = wwb.getSheet(0);
        ws.addCell(new Label(0, 0, "年份"));
        ws.addCell(new Label(1, 0, "豆瓣评价数"));
        ws.addCell(new Label(2, 0, "豆瓣得分"));
        ws.addCell(new Label(3, 0, "当当评价数"));
        ws.addCell(new Label(4, 0, "当当得分"));
        ws.addCell(new Label(5, 0, "亚马逊评价数"));
        ws.addCell(new Label(6, 0, "亚马逊得分"));
        ws.addCell(new Label(7, 0, "学术评论数"));

        this.closeSheet(wwb);

        Integer doubanHasComment = 0;
        Integer dangdangHasComment = 0;
        Integer amazonHasComment = 0;
        Integer scholarHasComment = 0;
        Double doubanAllYearCount = 0.0;
        Double doubanAllYearScore = 0.0;
        Double dangdangAllYearCount = 0.0;
        Double amazonAllYearCount = 0.0;
        Double amazonAllYearScore = 0.0;
        Double scholarAllYearCount = 0.0;

        //这是立项时间的起止，这里2004应该是不变的了，截止时间现在是写死的
        int start = 2004;
        int end = 2013;
        for (int i = start; i <= end; i++) {
            Double doubanTotalCount = 0.0;
            Double doubanTotalScore = 0.0;
            Double dangdangTotalCount = 0.0;
            Double amazonTotalCount = 0.0;
            Double amazonTotalScore = 0.0;
            Double scholarTotalCount = 0.0;
            Double newspapeTotalCount = 0.0;
            Integer doubanHasCommentYear = 0;
            Integer dangdangHasCommentYear = 0;
            Integer amazonHasCommentYear = 0;
            Integer scholarHasCommentYear = 0;
            for (int j = 0; j < rows - 1; j++) {
                if (Integer.parseInt(projectTime[j]) == i) {
                    if (!doubanCount[j].equals("--")) {
                        doubanTotalCount += Double.parseDouble(doubanCount[j]);
                        doubanHasCommentYear++;
                    }
                    if (!doubanScore[j].equals("--")) {
                        doubanTotalScore += Double.parseDouble(doubanScore[j]);
                    }
                    if (!dangdangCount[j].equals("--")) {
                        dangdangTotalCount += Double.parseDouble(dangdangCount[j]);
                        dangdangHasCommentYear++;
                    }
                    if (!amazonCount[j].equals("--")) {
                        amazonTotalCount += Double.parseDouble(amazonCount[j]);
                        amazonHasCommentYear++;
                    }
                    if (!amazonScore[j].equals("--")) {
                        amazonTotalScore += Double.parseDouble(amazonScore[j]);
                    }
                    if (!scholarCount[j].equals("0")) {
                        scholarTotalCount += Double.parseDouble(scholarCount[j]);
                        scholarHasCommentYear++;
                    }
                }
            }

            Double doubanAverageCount = doubanTotalCount / doubanHasCommentYear;
            Double doubanAverageScore = doubanTotalScore / doubanHasCommentYear;
            Double dangdangAverageCount = dangdangTotalCount / dangdangHasCommentYear;
            Double amazonAverageCount = amazonTotalCount / amazonHasCommentYear;
            Double amazonAverageScore = amazonTotalScore / amazonHasCommentYear;
            Double scholarAverageCount = scholarTotalCount / scholarHasCommentYear;

            doubanHasComment += doubanHasCommentYear;
            dangdangHasComment += dangdangHasCommentYear;
            amazonHasComment += amazonHasCommentYear;
            scholarHasComment += scholarHasCommentYear;

            doubanAllYearCount += doubanTotalCount;
            doubanAllYearScore += doubanTotalScore;
            dangdangAllYearCount += dangdangTotalCount;
            amazonAllYearCount += amazonTotalCount;
            amazonAllYearScore += amazonTotalScore;
            scholarAllYearCount += scholarTotalCount;

            wwb = this.writeSheet(outputPath);
            ws = wwb.getSheet(0);
            ws.addCell(new Number(0, i - start + 1, i));
            ws.addCell(new Number(1, i - start + 1, doubanAverageCount));
            ws.addCell(new Number(2, i - start + 1, doubanAverageScore));
            ws.addCell(new Number(3, i - start + 1, dangdangAverageCount));
            ws.addCell(new Number(4, i - start + 1, 5));
            ws.addCell(new Number(5, i - start + 1, amazonAverageCount));
            ws.addCell(new Number(6, i - start + 1, amazonAverageScore));
            ws.addCell(new Number(7, i - start + 1, scholarAverageCount));
            this.closeSheet(wwb);
        }

        wwb = this.writeSheet(outputPath);
        ws = wwb.getSheet(0);
        ws.addCell(new Label(0, end - start + 2, "总评论数"));
        ws.addCell(new Number(1, end - start + 2, doubanAllYearCount));
        ws.addCell(new Number(3, end - start + 2, dangdangAllYearCount));
        ws.addCell(new Number(5, end - start + 2, amazonAllYearCount));
        ws.addCell(new Number(7, end - start + 2, scholarAllYearCount));

        ws.addCell(new Label(0, end - start + 3, "有评论图书个数"));
        ws.addCell(new Number(2, end - start + 3, doubanHasComment));
        ws.addCell(new Number(4, end - start + 3, dangdangHasComment));
        ws.addCell(new Number(6, end - start + 3, amazonHasComment));
        ws.addCell(new Number(7, end - start + 3, scholarHasComment));

        ws.addCell(new Label(0, end - start + 4, "平均值"));
        ws.addCell(new Number(1, end - start + 4, doubanAllYearCount / doubanHasComment));
        ws.addCell(new Number(2, end - start + 4, doubanAllYearScore / doubanHasComment));
        ws.addCell(new Number(3, end - start + 4, dangdangAllYearCount / dangdangHasComment));
        ws.addCell(new Number(4, end - start + 4, 5));
        ws.addCell(new Number(5, end - start + 4, amazonAllYearCount / amazonHasComment));
        ws.addCell(new Number(6, end - start + 4, amazonAllYearScore / amazonHasComment));
        ws.addCell(new Number(7, end - start + 4, scholarAllYearCount / scholarHasComment));

        this.closeSheet(wwb);
    }


    public void 统计各种被引情况(String inputPath, String outputPath) throws IOException, BiffException, WriteException {
        Sheet sheet = this.readSheet(inputPath, 0);
        Integer rows = sheet.getRows();
        String[] projectTime = new String[rows - 1];
        String[] magazines = new String[rows - 1];
        String[] masters = new String[rows - 1];
        String[] doctors = new String[rows - 1];
        String[] conferences = new String[rows - 1];
        String[] selfCitations = new String[rows - 1];
        String[] selfInstituteCitations = new String[rows - 1];

        for (int i = 1; i < rows; i++) {
            projectTime[i - 1] = sheet.getCell(CellName.projectTime.getValue(), i).getContents();
            magazines[i - 1] = sheet.getCell(CellName.magazine.getValue(), i).getContents();
            masters[i - 1] = sheet.getCell(CellName.master.getValue(), i).getContents();
            doctors[i - 1] = sheet.getCell(CellName.doctor.getValue(), i).getContents();
            conferences[i - 1] = sheet.getCell(CellName.conference.getValue(), i).getContents();
            selfCitations[i - 1] = sheet.getCell(CellName.authorSelf.getValue(), i).getContents();
            selfInstituteCitations[i - 1] = sheet.getCell(CellName.instituteSelf.getValue(), i).getContents();
        }

        WritableWorkbook wwb = this.writeSheet(outputPath);
        WritableSheet ws = wwb.getSheet(0);
        ws.addCell(new Label(0, 0, "立项时间"));
        ws.addCell(new Label(1, 0, "总被引"));
        ws.addCell(new Label(2, 0, "学位论文"));
        ws.addCell(new Label(3, 0, "期刊论文"));
        ws.addCell(new Label(4, 0, "会议论文"));
        ws.addCell(new Label(5, 0, "自引率"));
        ws.addCell(new Label(6, 0, "机构自引率"));
        this.closeSheet(wwb);


        //这是立项时间的起止，这里2004应该是不变的了，截止时间现在是写死的
        int start = 2004;
        int end = 2014;

        Double degreeAllTotal = 0.0;
        Double magazineAllTotal = 0.0;
        Double conferenceAllTotal = 0.0;
        Double selfCitAllTotal = 0.0;
        Double selfInstituteAllTotal = 0.0;
        Double allTotal = 0.0;
        for (int i = start; i <= end; i++) {
            Double degreeYearTotal = 0.0;
            Double magazineYearTotal = 0.0;
            Double conferenceYearTotal = 0.0;
            Double selfCitYearTotal = 0.0;
            Double selfInstituteYearTotal = 0.0;
            Double yearTotal = 0.0;
            for (int j = 0; j < rows - 1; j++) {
                if(magazines[j].equals("")){
                    continue;
                }
                if (Integer.parseInt(projectTime[j]) == i) {
                    if (!magazines[j].equals("--") && magazines[j] != null) {
                        magazineYearTotal += Double.parseDouble(magazines[j]);
                    }
                    if (!conferences[j].equals("--") && conferences[j] != null) {
                        conferenceYearTotal += Double.parseDouble(conferences[j]);
                    }
                    if (!doctors[j].equals("--") && doctors[j] != null) {
                        degreeYearTotal += Double.parseDouble(doctors[j]);
                    }
                    if (!masters[j].equals("--") && masters[j] != null) {
                        degreeYearTotal += Double.parseDouble(masters[j]);
                    }
                    if (!selfCitations[j].equals("--") && selfCitations[j] != null) {
                        selfCitYearTotal += Double.parseDouble(selfCitations[j]);
                    }
                    if (!selfInstituteCitations[j].equals("--") && selfInstituteCitations[j] != null) {
                        selfInstituteYearTotal += Double.parseDouble(selfInstituteCitations[j]);
                    }
                }
            }
            yearTotal = degreeYearTotal + conferenceYearTotal + magazineYearTotal;

            allTotal += yearTotal;
            degreeAllTotal += degreeYearTotal;
            conferenceAllTotal += conferenceYearTotal;
            magazineAllTotal += magazineYearTotal;
            selfCitAllTotal += selfCitYearTotal;
            selfInstituteAllTotal += selfInstituteYearTotal;

            wwb = this.writeSheet(outputPath);
            ws = wwb.getSheet(0);
            ws.addCell(new Number(0, i - start + 1, i));
            ws.addCell(new Number(1, i - start + 1, yearTotal));
            ws.addCell(new Number(2, i - start + 1, degreeYearTotal));
            ws.addCell(new Number(3, i - start + 1, magazineYearTotal));
            ws.addCell(new Number(4, i - start + 1, conferenceYearTotal));
            ws.addCell(new Number(5, i - start + 1, selfCitYearTotal / yearTotal));
            ws.addCell(new Number(6, i - start + 1, selfInstituteYearTotal / yearTotal));
            this.closeSheet(wwb);
        }

        wwb = this.writeSheet(outputPath);
        ws = wwb.getSheet(0);
        ws.addCell(new Label(0, end - start + 2, "总计"));
        ws.addCell(new Number(1, end - start + 2, allTotal));
        ws.addCell(new Number(2, end - start + 2, degreeAllTotal));
        ws.addCell(new Number(3, end - start + 2, magazineAllTotal));
        ws.addCell(new Number(4, end - start + 2, conferenceAllTotal));
        ws.addCell(new Number(5, end - start + 2, selfCitAllTotal / allTotal));
        ws.addCell(new Number(6, end - start + 2, selfInstituteAllTotal / allTotal));
        this.closeSheet(wwb);

    }

    /**
     * 通用读表
     *
     * @param path
     * @return
     */
    public Sheet readSheet(String path, int sheetNum) throws IOException, BiffException {
        InputStream instream = new FileInputStream(path);
        Workbook readwb = Workbook.getWorkbook(instream);
        Sheet sheet = readwb.getSheet(sheetNum);
        return sheet;
    }

    /**
     * 通用写表
     *
     * @param path
     * @return
     */
    public WritableWorkbook writeSheet(String path) throws IOException, BiffException, WriteException {
        InputStream instream = new FileInputStream(path);
        Workbook rwb = Workbook.getWorkbook(instream);
        WritableWorkbook wwb = Workbook.createWorkbook(new File(path), rwb);//copy
        return wwb;

    }

    /**
     * 通用关闭表格
     *
     * @param wwb
     * @throws IOException
     * @throws WriteException
     */
    public void closeSheet(WritableWorkbook wwb) throws IOException, WriteException {
        wwb.write();
        wwb.close();
    }


    public static void main(String[] args) throws Exception {
        TableForReport report = new TableForReport();
//        report.计算峰值间隔("resources/result.xls");
//        report.计算首次被引时间("resources/result.xls");
//        report.计算年均引用("resources/result.xls");
//      其他几个表格数据计算可以明年再补充了
//        report.统计各种评论数("resources/result.xls", "resources/process.xls");
        report.统计各种被引情况("resources/result.xls", "resources/process.xls");
    }
}
