public class Split {
    public static void main(String[] args) {
        String str = "5780｜QC";
        System.out.println(str.contains("If"));
        String[] a = str.split("｜");
        System.out.println(a);

        System.out.println(SplitString(str));
        String c = "*V";
        System.out.println("TEST "+c.substring(0,1));
        System.out.println(c.substring(0,1).equals("*"));
        System.out.println((c.contains("*")));

    }



    private static String SplitString(String ProductNameValue)
    {
        String[] value =  ProductNameValue.split("｜");
        return value[1];
    }
}
