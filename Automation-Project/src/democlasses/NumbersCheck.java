package democlasses;

public class NumbersCheck {
	
	public static void main(String []args){
		
		int num [] = {9,5,1,5,9,7,12};
		boolean chk = false;
		int c=0;
		for(int i=0; i<num.length-1; i++){
			
			int first = num[i];
			int scnd = num[i+1];
			int d = 0;
			
			if(first>scnd){
				d = first-scnd;
			}else if(scnd>first){
				d = scnd-first;
			}
			
			System.out.println("value "+c+" "+d);
			if(i!=0){
				if(c==d){
					chk = true;
				}else {
					chk=false;
					break;
				}
			}
			c=d;
		}
		
		System.out.println(chk);
	}

}
