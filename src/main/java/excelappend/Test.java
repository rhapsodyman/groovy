package excelappend;

public class Test {
	static String abc = "5";
	
	public static void main(String[] args) {
		someFunction();
	}
	
	static void someFunction(){
		System.out.println(abc);
		
		String abc = "6";
		
		System.out.println(abc);
		
	}

}
