package com.demo.model;

@lombok.Getter
@lombok.Setter
@lombok.NoArgsConstructor

public class Person {

	public String Name;
	private String Gender;
	private String Age;
	private String Result;

	@Override
	public String toString(){
		return "name="+Name+", gender="+Gender +", age="+Age +"";
	}
}
