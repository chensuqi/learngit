
public class Test {
    public static void main(String[] args) {
        System.out.println(28D/10D);

        String Str = new String("www.google.com");

        System.out.print("匹配成功返回值 :" );
//        System.out.println(Str.replaceAll("(.*)google(.*)", "runoob" ));
        System.out.print("匹配失败返回值 :" );
        System.out.println(Str.replaceAll("(.*)taobao(.*)", "runoob" ));
    }
}
