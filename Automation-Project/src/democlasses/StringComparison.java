package democlasses;

public class StringComparison {
		
	public static void main(String []args){
		
		String s1 ="Hello";
		String s2 = "Hello";
		String s3 = new String("Hello");
		
		if(s1.equals(s2)){
			System.out.println("match");
		}else {
			System.out.println("not match");
		}
		
		if(s1.equals(s3)){
			System.out.println("match");
		}else {
			System.out.println("not match");
		}
		
	}
	
}
