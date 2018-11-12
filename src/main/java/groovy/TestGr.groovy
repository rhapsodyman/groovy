package groovy

import org.apache.commons.lang3.StringUtils

import excelappend.ExamDTO


// https://itexpertsconsultant.wordpress.com/2017/02/22/create-a-new-groovy-project-using-maven/
//https://learnxinyminutes.com/docs/groovy/

class TestGr {

	static main(args) {
		
		def math = new ExamDTO();
		math.setDiscipline("Math");
		math.setRoom("B1");
		
		def biol = new ExamDTO();
		biol.setDiscipline("Biology");
		biol.setRoom("B2");
		
		def exams = []
		exams.add(math)
		exams.add(biol)
		
//		exams.each { println "Discipline is  + $it" }
		for (ex in exams) {
			println ex.getProperties().get("discipline")
			println ex.getProperties().get("room")
		}
		
		print "\n\n\n"
		
		
//		using third party java lib
		def name = "Andrei"
		print "Upper case name is " + StringUtils.upperCase(name) + "\n"
		print "Hello"

		def technologies = []

		/*** Adding a elements to the list ***/

		// As with Java
		technologies.add("Grails")

		// Left shift adds, and returns the list
		technologies << "Groovy"

		// Add multiple elements
		technologies.addAll(["Gradle", "Griffon"])

		/*** Removing elements from the list ***/

		// As with Java
		technologies.remove("Griffon")

		// Subtraction works also
		technologies = technologies - 'Grails'

		/*** Iterating Lists ***/

		// Iterate over elements of a list
		technologies.each { println "Technology: $it"}
		technologies.eachWithIndex { it, i -> println "$i: $it"}
	}

}
