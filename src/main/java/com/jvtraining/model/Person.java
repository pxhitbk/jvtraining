/**
 * <p>
 * </p>
 *
 * @author hungpx
 * @since
 */
package com.jvtraining.model;

/**
 * @author august
 *
 */
public class Person {
	private Long id;
	private String name;
	private Gender gender;
	private String carrer;

	public Person() {
	}

	public Person(Long id, String name, Gender sex, String carrer) {
		this.id = id;
		this.name = name;
		this.gender = sex;
		this.carrer = carrer;
	}

	public Long getId() {
		return id;
	}

	public void setId(Long id) {
		this.id = id;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public Gender getGender() {
		return gender;
	}

	public void setGender(Gender gender) {
		this.gender = gender;
	}

	public String getCarrer() {
		return carrer;
	}

	public void setCarrer(String carrer) {
		this.carrer = carrer;
	}

	/*
	 * (non-Javadoc)
	 * 
	 * @see java.lang.Object#toString()
	 */
	@Override
	public String toString() {
		return "[" + hashCode() + "], id:" + id + ", name:" + name;
	}

	/*
	 * (non-Javadoc)
	 * 
	 * @see java.lang.Object#equals(java.lang.Object)
	 */
	@Override
	public boolean equals(Object obj) {
		if (obj == null)
			return false;
		if (!(obj instanceof Person))
			return false;
		Person p = (Person) obj;
		if (p.getId() == null)
			return false;
		if (p.getId().equals(id))
			return true;
		return false;
	}

	@Override
	public Person clone() throws CloneNotSupportedException {
		// Person p = new Person();
		// p.setCarrer(carrer);
		// p.setSex(sex);
		// p.setId(id);
		// p.setName(name);
		// p.setCarrer(carrer);
		return (Person) super.clone();
	}

}
