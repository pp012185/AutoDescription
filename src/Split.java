public class Split {
    public static void main(String[] args) {
        String str = "If,Kind,N-CH MOSFET=>EX20.1,P-CH MOSFET=>EX20.2";
        System.out.println(str.contains("If"));
        String[] a = str.split(",");
//        System.out.println(a);
        for (int i = 0 ; i<a.length;i++){
            System.out.println(a[i]);

        }
        String[] b = a[2].split("=>");
        String b1 = b[0];
        String b2 = b[1];
        System.out.println(b1);
        System.out.println(b2);

    }
}
