package spider;

import org.apache.http.HttpEntity;
import org.apache.http.HttpResponse;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.impl.client.DefaultHttpClient;
import org.apache.http.util.EntityUtils;

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import pojo.CnkiResult;
import pojo.GTResult;
import pojo.SearchResult;

import java.util.*;

import java.io.IOException;
import java.net.URLEncoder;
import java.util.regex.Matcher;
import java.util.regex.Pattern;


/**
 * Created by Administrator on 2015/9/9.
 */
public class NlcSpider {
    private String cookieStr;
    private boolean authorBool = false;
    private boolean pubBool = false;
    private boolean titleBool = false;


    public String getParams(SearchResult searchResult) throws IOException {
        List<String> titles = new ArrayList<String>();
        List<String> authors = new ArrayList<String>();
        List<String> pubs = new ArrayList<String>();
        String[] format = new Util().formatSearchResult(searchResult);
        String searchTitle = format[0];
        String title = "";
        String author = "";
        String pub = "";
        String titleBak = format[3];
        String param = "http://find.nlc.cn/search/showDocDetails?docId=";
        System.out.println("_______________________");
        String url = "http://find.nlc.cn/search/doSearch?query=" + URLEncoder.encode(searchTitle, "utf-8") + "&secQuery=&actualQuery=" + URLEncoder.encode(searchTitle, "utf-8") + "mediatype%3A(0%20OR%201%20OR%202)%20&searchType=2&docType=%E5%9B%BE%E4%B9%A6&mediaTypes=0,1,2&isGroup=isGroup&targetFieldLog=%E5%85%A8%E9%83%A8%E5%AD%97%E6%AE%B5&orderBy=RELATIVE";
        String resText = getEntity(url);
//        System.out.println(resText);

        Pattern patternTitle = Pattern.compile("<a id=\"([^\\<]*[istitle=\"true\"]+[^\\<]*?)<([span/a]*)([^\\<]*)");
        Matcher matcherTitle = patternTitle.matcher(resText);
        while (matcherTitle.find()) {
            titleBak = "";
            title = matcherTitle.group();
            if (title.contains("isTitle=\"true\"")) {
                //参数
                param += title.trim().substring(7, title.indexOf("\"", 8));
                param += "&dataSource=ucs01&query=";
                //书名
                if (title.contains("<span")) {
                    title = title.substring(title.lastIndexOf("tle=") + 5, title.lastIndexOf(">") - 1);
//                    System.out.println(title);
                } else {
                    title = title.substring(title.indexOf("ank\">") + 5, title.lastIndexOf("<")).trim();
//                    System.out.println(title);
                }
            }
            titleBool = title.contains(searchTitle) || searchTitle.contains(title);
            if (titleBool) {
                titles.add(title);
                System.out.print("要查的书名：" + searchTitle + "   检索到的：" + title + "   ");
                Pattern patternPub = Pattern.compile("<span class=\"pub\">出版社([^\\>]*)<([^\\<]*)");
                Matcher matcherPub = patternPub.matcher(resText);
                while (matcherPub.find()) {
                    pub = matcherPub.group();
                    pub = pub.substring(pub.lastIndexOf(">") + 1, pub.length()).trim();
                    pubBool = pub.contains(searchResult.getPublisher()) || searchResult.getPublisher().contains(pub);
                    if (pubBool) {
                        pubs.add(pub);
                        System.out.println(pub + "   ");
                        param += URLEncoder.encode(title, "utf-8");
                        System.out.println(param);
                        //作者 不强制
                        Pattern patternAuthor = Pattern.compile("<p>著者([^\\>]*)<([^\\<]*)");
                        Matcher matcherAuthor = patternAuthor.matcher(resText);
                        while (matcherAuthor.find()) {
                            author = matcherAuthor.group();
                            author = author.substring(author.lastIndexOf(">") + 1, author.length()).trim();
                            authorBool = author.contains(searchResult.getAuthor()) || searchResult.getAuthor().contains(author);
                            if (authorBool) {
                                authors.add(author);
                                System.out.print(author + "  作者匹配成功   ");
                                break;
                            }
                        }
                        if (authors.size() <= 0) {
                            System.out.println(searchResult.getAuthor() + "   没找到作者------");
                        }
                        return param;
                    }
                }
                if (pubs.size() <= 0) {
                    System.out.println(searchResult.getAuthor() + "   书名+出版社匹配失败------");
                }
            }
        }
        if (titles.size() <= 0) {
            System.out.println(searchTitle + "   没有找到书名，尝试备用书名------");
        }
        if (!Objects.equals(titleBak, "")) {
            Pattern patternTitle1 = Pattern.compile("<a id=\"([^\\<]*[istitle=\"true\"]+[^\\<]*?)<([span/a]*)([^\\<]*)");
            Matcher matcherTitle1 = patternTitle1.matcher(resText);
            while (matcherTitle1.find()) {
                title = matcherTitle1.group();
                if (title.contains("isTitle=\"true\"")) {
                    //参数
                    param += title.trim().substring(7, title.indexOf("\"", 8));
                    param += "&dataSource=ucs01&query=";
                    //书名
                    if (title.contains("<span")) {
                        title = title.substring(title.lastIndexOf("tle=") + 5, title.lastIndexOf(">") - 1);
//                    System.out.println(title);
                    } else {
                        title = title.substring(title.indexOf("ank\">") + 5, title.lastIndexOf("<")).trim();
//                    System.out.println(title);
                    }
                }
                titleBool = titleBak.contains(title) || title.contains(titleBak);
                System.out.println("要查的书名：" + titleBak + "   检索到的：" + title + "   ");
                if (titleBool) {
                    titles.clear();
                    titles.add(titleBak);
                    Pattern patternPub = Pattern.compile("<span class=\"pub\">出版社([^\\>]*)<([^\\<]*)");
                    Matcher matcherPub = patternPub.matcher(resText);
                    while (matcherPub.find()) {
                        pub = matcherPub.group();
                        pub = pub.substring(pub.lastIndexOf(">") + 1, pub.length()).trim();
                        pubBool = pub.contains(searchResult.getPublisher()) || searchResult.getPublisher().contains(pub);
                        if (pubBool) {
                            pubs.add(pub);
                            System.out.println(pub + "   ");
                            param += URLEncoder.encode(titleBak, "utf-8");
                            System.out.println(param);
                            //作者 不强制
                            Pattern patternAuthor = Pattern.compile("<p>著者([^\\>]*)<([^\\<]*)");
                            Matcher matcherAuthor = patternAuthor.matcher(resText);
                            while (matcherAuthor.find()) {
                                author = matcherAuthor.group();
                                author = author.substring(author.lastIndexOf(">") + 1, author.length()).trim();
                                authorBool = author.contains(searchResult.getAuthor()) || searchResult.getAuthor().contains(author);
                                if (authorBool) {
                                    authors.add(author);
                                    System.out.println("  作者匹配成功:" + author);
                                    break;
                                }
                            }
                            if (authors.size() <= 0) {
                                System.out.println(searchResult.getAuthor() + "   没找到作者------");
                            }
                            return param;
                        }
                    }
                    if (pubs.size() <= 0) {
                        System.out.println(searchResult.getAuthor() + "   书名+出版社匹配失败------");
                    }
                }
            }
            if (titles.size() <= 0) {
                System.out.println(searchResult.getTitle() + "   这回真没有找到书名------");
            }
        }
        return null;
    }

    public String getEntity(String url) throws IOException {
        CnkiResult cnkiResult = new CnkiResult();
        cookieStr = "Hm_lvt_7fb5be3a55ebab6d1f7bf2e54d83e203=1441642372,1441728317; Hm_lpvt_7fb5be3a55ebab6d1f7bf2e54d83e203=1441728317";
        DefaultHttpClient httpClient = new DefaultHttpClient();
        HttpGet httpGet = new HttpGet(url);
        HttpResponse response = null;
        HttpEntity entity = null;
        Map<String, String> headers = new HashMap<String, String>();
        httpGet.setHeader("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8");
        httpGet.setHeader("User-Agent", "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/53.0.2785.89 Safari/537.36");
        httpGet.setHeader("Accept-Language", "zh-CN,zh;q=0.8");
        httpGet.setHeader("Cookie", cookieStr);

        httpGet.setHeader("Host", "find.nlc.cn");
        response = httpClient.execute(httpGet);
        entity = response.getEntity();
        String resText = EntityUtils.toString(entity, "utf-8");
        return resText;
    }

    public GTResult getTable(String url) throws IOException {
        GTResult gtResult = new GTResult();
        String resText = getEntity(url);
//        System.out.println(resText);
        //概述信息
        Document doc = Jsoup.parse(resText);
        Elements titles = doc.select("div h4 b");
        for (Element b : titles) {
            System.out.println(b.html());
            gtResult.setTitle(b.html().trim());
        }
        Elements info = doc.select("div.info p");
        for (Element p : info) {

            if (p.html().contains("文献类型")) {
                String type = p.select("em").html().trim();
                System.out.println("文献类型: " + type);
                gtResult.setBooktype(type);
            } else if (p.html().contains("责任者") && !p.html().contains("所有责任者")) {
                String author = p.select("a").html().trim();
                System.out.println("责任者: " + author);
                gtResult.setMainAuthor(author);
            } else if (p.html().contains("出版、发行者")) {
                String publiser = p.select("a").html().trim();
                System.out.println("出版、发行者: " + publiser);
                gtResult.setPublisher(publiser);
            } else if (p.html().contains("出版发行时间")) {
                String pubtime = p.select("em").html().trim();
                System.out.println("出版发行时间: " + pubtime);
                gtResult.setPubTime(pubtime);
            }
        }

        int i = 0;
        Elements detailInfo = doc.select("div#detail-info p");
        for (Element p : detailInfo) {
            i++;
            String commonStr = p.ownText();
            if (commonStr.contains("：")) {
                commonStr = commonStr.substring(commonStr.indexOf("：") + 1, commonStr.length());
            }
            if (p.select("b").get(0).text().contains("所有题名")) {
                String allTitle = commonStr;
                System.out.println("所有题名: " + allTitle);
                gtResult.setAllTitle(allTitle);
            } else if (p.select("b").get(0).text().contains("标识号")) {
                String isbn = commonStr;
                isbn = isbn.replaceAll("&nbsp;", "").trim();
                gtResult.setISBN(isbn);
                System.out.println("标识号---- " + isbn);
            } else if (p.select("b").get(0).text().contains("所有责任者")) {
                String allAuthor = commonStr;
                System.out.println("所有责任者---- " + allAuthor);
                gtResult.setAllAuthor(allAuthor);
            } else if (p.select("b").get(0).text().contains("出版、发行地")) {
                String publishPlace = commonStr;
                System.out.println("出版、发行地---- " + publishPlace);
                gtResult.setPublishPlace(publishPlace);
            } else if (p.select("b").get(0).text().contains("题名、责任者附注")) {
                String otherAuthor = commonStr;
                gtResult.setOtherAuthor(otherAuthor);
                System.out.println("题名、责任者附注---- " + otherAuthor);
            } else if (p.select("b").get(0).text().contains("语种")) {
                String language = commonStr;
                System.out.println("语种---- " + language);
                gtResult.setLangguage(language);
            } else if (p.select("b").get(0).text().contains("分类")) {
                String sort = commonStr.substring(commonStr.indexOf(":")+1, commonStr.length());
                System.out.println("分类---- " + sort);
                gtResult.setLibrarySort(sort);
            } else if (p.select("b").get(0).text().contains("载体形态")) {
                String page = commonStr;
                System.out.println("载体形态---- " + page);
                gtResult.setPage(page);
            } else if (p.select("b").get(0).text().contains("关键词")) {

            }
// else if (p.select("b").get(0).text().contains("丛编题名")) {
//                String page = commonStr.substring(commonStr.indexOf(":"), commonStr.length());
//                System.out.println("载体形态---- " + page);
//                gtResult.setPage(page);
//            }
            else {
                System.out.print("   -----有奇葩属性-----:" + p.html());
            }
        }

        //详细信息
//        Pattern patternDetailPart = Pattern.compile("<p>(\\s*)<script>([^\\<]*)</script>(\\s*)<b>([^\\<]*)</b>([^\\<]*)</p>");
//        Matcher matcherDetailPart = patternDetailPart.matcher(resText);
//        int i = 0;
//        while (matcherDetailPart.find()) {
//            i++;
//            String str = matcherDetailPart.group();
//            String param;
//            param = str.substring(str.indexOf("<b>") + 3, str.indexOf("</b>")).trim();
//            str = str.substring(str.indexOf("</b>") + 5, str.lastIndexOf("<")).trim();
//            if (param.equals("所有责任者")) {
//                str = str.replaceAll("&nbsp;", "").trim();
//                gtResult.setAllAuthor("     " + str);
//
//            } else if (param.equals("所有题名")) {
//
//            } else if (param.equals("标识号")) {
//                str = str.replaceAll("&nbsp;", "").trim();
//                gtResult.setISBN(str);
//                System.out.println("   " + str);
//
//            } else if (param.equals("出版、发行地")) {
//                gtResult.setPublishPlace(str);
//                System.out.println("   " + str);
//
//            } else if (param.equals("题名、责任者附注")) {
//                if (str.contains(":")) {
//                    str = str.substring(str.indexOf(":"), str.length());
//                    gtResult.setOtherAuthor(str);
//                    System.out.println("   " + str);
//                } else if (str.contains("：")) {
//                    str = str.substring(str.indexOf("："), str.length());
//                    gtResult.setOtherAuthor(str);
//                    System.out.println("   " + str);
//
//                }
//            } else if (param.equals("语种")) {
//                gtResult.setLangguage(str);
//                System.out.println("   " + str);
//
//            } else if (param.equals("分类")) {
//                str = str.replaceAll("&nbsp;", "").trim();
//                gtResult.setLibrarySort(str);
//                System.out.println("   " + str);
//
//            } else if (param.equals("载体形态")) {
//                gtResult.setPage(str);
//                System.out.println("   " + str);
//
//            } else if (param.equals("关键词")) {
//                System.out.println("    --有关键词：" + str);
//
//            } else if (param.equals("丛编题名")) {
//                System.out.println("    --有丛编题名：" + str);
//
//            } else {
//                System.out.print("   -----有奇葩属性-----");
//            }
//        }
        if (i != 7 && i != 9) {
            System.out.println("有不是7条或者9条的信息");
        }
        return gtResult;
    }
}
