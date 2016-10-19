package mains;

import excel.ExcelProcess;
import pojo.GTResult;
import pojo.SearchResult;
import spider.DangdangSpider;
import spider.NlcSpider;

/**
 * Created by Administrator on 2015/8/18.
 */
public class MainBookInfo {
    //当当上获取图书信息
    public void dangdangBookInfo() throws Exception {
        ExcelProcess reader = new ExcelProcess();
        SearchResult[] searchResults = reader.reader("resources/source-init.xls");
        DangdangSpider dangdangSpider = new DangdangSpider();
        //拉取书目详情url，这个跟获取评价逻辑类似，之前的叫getPrams，返回一个字符串，已经删掉了
        String url = "";
        for (int i = 0; i < searchResults.length; i++) {
            if (!url.equals("")) {
                Thread.sleep(2000);
                GTResult gtResult = dangdangSpider.getBookInfo(url);
                reader.writeDangdangBookInfo(gtResult);
            } else {
                reader.writeDangdangBookInfo(null);
            }
            Thread.sleep(2000);
        }
    }

    //国图的图书信息
    public void nlcBookInfo() throws Exception {
        ExcelProcess reader = new ExcelProcess();
        SearchResult[] searchResults = reader.reader("resources/source-init.xls");
        GTResult[] gtResults = new GTResult[searchResults.length];
        NlcSpider nlcSpider = new NlcSpider();
        String url;
        for (int i = 0; i < searchResults.length; i++) {
            url = nlcSpider.getParams(searchResults[i]);
            gtResults[i] = new GTResult();
            if (url != null) {
                gtResults[i] = nlcSpider.getTable(url);
            } else {
                gtResults[i] = null;
            }
            Thread.sleep(2000);
        }
        reader.writeNlc(gtResults);

//        SearchResult searchResult = new SearchResult();
//        searchResult.setTitle("大国崛起制高点");
//        searchResult.setPublisher("人民出版社");
//        searchResult.setAuthor("胡雪梅");
//        url = nlcSpider.getParams(searchResult);
    }

    public static void main(String[] args) throws Exception {
        MainBookInfo bookInfo = new MainBookInfo();
        bookInfo.nlcBookInfo();
        bookInfo.dangdangBookInfo();
    }
}
