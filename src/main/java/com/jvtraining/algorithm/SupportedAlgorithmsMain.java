/**
 * <p>
 * </p>
 *
 * @author hungpx
 * @since
 */
package com.jvtraining.algorithm;

import java.util.Collections;
import java.util.Comparator;
import java.util.List;

import com.jvtraining.model.Person;
import com.jvtraining.service.PersonService;

/**
 * @author august
 *
 */
public class SupportedAlgorithmsMain {
	static PersonService personService = new PersonService();

	public static void main(String[] args) {
		List<Person> ps = personService.getAll();
		Collections.sort(ps, new PersonNameCompare());

		Person tom = new Person();
		tom.setName("Tom");
		int pos = Collections.binarySearch(ps, tom, new PersonNameCompare());

		System.out.println("Tom stay at: " + pos);

		for (Person p : ps) {
			System.out.println(p.toString());
		}
	}

	static class PersonNameCompare implements Comparator<Person> {

		/*
		 * (non-Javadoc)
		 *
		 * @see java.util.Comparator#compare(java.lang.Object, java.lang.Object)
		 */
		public int compare(Person o1, Person o2) {
			return o1.getName().compareTo(o2.getName());
		}
	}
}
