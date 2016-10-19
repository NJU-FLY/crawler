package excel;

/**
 * Created by Administrator on 2016/9/19.
 * 表格名称对应的列号
 */
public enum CellName {

    title(0),
    author(1),
    publisher(2),
    pubTime(3),
    cit2016(4),
    cit2015(5),
    cit2014(6),
    cit2013(7),
    cit2012(8),
    cit2011(9),
    cit2010(10),
    cit2009(11),
    cit2008(12),
    cit2007(13),
    leftTotal(14),
    crawledTotal(15),
    rightTotal(16),
    magazine(17),
    master(18),
    doctor(19),
    conference(20),
    authorSelf(21),
    instituteSelf(22),
    projectTime(23),
    doubanCommentCount(24),
    doubanScore(25),
    dangdangCommentCount(26),
    dangdangScore(27),
    amazonCommentCount(28),
    amazonScore(29),
    newspaperCommentCount(30),
    scholarCommentCount(31),

    年均引用(37),
    首次被引时间(39),
    峰值间隔(41),

    allAuthor(43),
    isbn(44),
    otherAuthor(46),
    language(47),
    librarySort(48),
    pages(49),
    price(50),
    bookType(52);


    private int value;

    CellName(int value) {
        this.value = value;
    }

    public int getValue(){
        return value;
    }

}
