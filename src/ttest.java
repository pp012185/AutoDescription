public class ttest {
    public static void main(String[] args) {
        String a = "Process Type!";
        String b = "b b";
        String c = "c c  c   c";
        //a=a.replaceAll("!","");
        b=b.replaceAll("\\s+","");
        c=c.replaceAll("\\s+", "");
        System.out.println("a:"+a+",");
        System.out.println("b:"+b+",");
        System.out.println("c:"+c+",");
        System.out.println(a.equals("Process Type") || a.equals("Process Type!"));

    }

    private static int generateAccordingToString(String a) {

        return 0;
    }
}
