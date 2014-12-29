/**
 * <p>
 * </p>
 *
 * @author hungpx
 * @since
 */
package com.jvtraining.service;

import java.util.ArrayList;
import java.util.List;

import com.jvtraining.model.Person;
import com.jvtraining.model.Sex;

/**
 * @author august
 *
 */
public class PersonService {
	public List<Person> getAll() {
		List<Person> ps = new ArrayList<Person>();
		String[] names = { "Tom", "Jerry", "Donal", "Scoopy Doo", "Mickey" };
		Sex[] gender = { Sex.MALE, Sex.FEMALE };
		String[] carrer = { "Dev", "Engineer", "Doctor", "Software engineer", "Musican" };
		for (int i = 0; i < names.length; i++) {
			Person p = new Person((long) i, names[i], gender[0], carrer[i]);
			ps.add(p);
		}
		return ps;
	}
}
