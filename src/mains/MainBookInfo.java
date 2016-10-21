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
        GTResult[] gtResults = new GTResult[searchResults.length];
        String url = "";
        for (int i = 0; i < searchResults.length; i++) {
            url = dangdangSpider.getBookInfoUrl(searchResults[i]);
            if (!url.equals("")) {
                gtResults[i] = dangdangSpider.getBookInfo(url);
            } else {
                gtResults[i] = null;
            }
            Thread.sleep(2000);
        }
        reader.writeDangdangBookInfo(gtResults);

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
//        bookInfo.nlcBookInfo();
        bookInfo.dangdangBookInfo();
    }
}
