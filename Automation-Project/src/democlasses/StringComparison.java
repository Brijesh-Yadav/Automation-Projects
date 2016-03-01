package democlasses;

public class StringComparison {
		
	public static void main(String []args){
		
		String s1 ="Hello";
		String s2 = "Hello";
		String s3 = new String("Hello");
		String s4 = new String("Hello");
		
		System.out.println(s1+ "");
		
		
		System.out.println(s1+ "");
		
		System.out.println(s1+ "");
		
		/*
		if(s1.equals(s2)){
			System.out.println("match");
		}else {
			System.out.println("not match");
		}
		*/
		
		if(s1.equals(s3)){
			System.out.println("match");
		}else {
			System.out.println("not match");
		}
		
	}
	
}
