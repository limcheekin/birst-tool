/**
 * 
 */
package com.vobject.birst.tool

/**
 * @author limcheek
 * Copy from GrailsNameUtils
 * https://github.com/grails/grails-core/blob/master/grails-bootstrap/src/main/groovy/grails/util/GrailsNameUtils.java
 *
 */


class NameUtils {

	/**
	 * Modified method.
	 * Returns the property name representation of the given name.
	 *
	 * @param name The name to convert
	 * @return The property name representation
	 */
	public static String getPropertyName(String name) {
		String propertyName
		
		// Strip any package from the name.
		int pos = name.lastIndexOf('.');
		if (pos != -1) {
			name = name.substring(pos + 1);
		}
		
		// Check whether the name begins with two upper case letters.
		if (name.length() > 1 && Character.isUpperCase(name.charAt(0)) &&
			Character.isUpperCase(name.charAt(1))) {
			propertyName = name
		} else {
			propertyName = name.substring(0,1).toLowerCase(Locale.ENGLISH) + name.substring(1);
		}

		if (propertyName.indexOf(' ') > -1) {
			propertyName = propertyName.replaceAll("\\s", "");
		}

		return propertyName;
	}

	/**
	 * Returns the class name without the package prefix.
	 *
	 * @param className The class name to get a short name for
	 * @return The short name of the class
	 */
	public static String getShortName(String className) {
		int i = className.lastIndexOf(".");
		if (i > -1) {
			className = className.substring(i + 1, className.length());
		}
		return className;
	}


	/**
	 * Converts a property name into its natural language equivalent eg ('firstName' becomes 'First Name')
	 * @param name The property name to convert
	 * @return The converted property name
	 */
	public static String getNaturalName(String name) {
		name = getShortName(name);
		List<String> words = new ArrayList<String>();
		int i = 0;
		char[] chars = name.toCharArray();
		for (int j = 0; j < chars.length; j++) {
			char c = chars[j];
			String w;
			if (i >= words.size()) {
				w = "";
				words.add(i, w);
			}
			else {
				w = words.get(i);
			}
			if (Character.isLowerCase(c) || Character.isDigit(c)) {
				if (Character.isLowerCase(c) && w.length() == 0) {
					c = Character.toUpperCase(c);
				}
				else if (w.length() > 1 && Character.isUpperCase(w.charAt(w.length() - 1))) {
					w = "";
					words.add(++i,w);
				}
				words.set(i, w + c);
			}
			else if (Character.isUpperCase(c)) {
				if ((i == 0 && w.length() == 0) || (Character.isUpperCase(w.charAt(w.length() - 1)) && Character.isUpperCase(chars[j-1]))) {
					words.set(i, w + c);
				}
				else {
					words.add(++i, String.valueOf(c));
				}
			}
		}
		StringBuilder buf = new StringBuilder();
		for (Iterator<String> j = words.iterator(); j.hasNext();) {
			String word = j.next();
			buf.append(word);
			if (j.hasNext()) {
				buf.append(' ');
			}
		}
		return buf.toString();
	}

}
