print("=== Student Grade System ===")

student_name = input("Enter student name: ")
number_of_subjects = int(input("Enter number of subjects: "))
while number_of_subjects <= 0:
    print("Number of subjects must be at least 1.")
    number_of_subjects = int(input("Enter number of subjects: "))

total_marks = 0

for subject_number in range(1, number_of_subjects + 1):
    mark = float(input(f"Enter marks for subject {subject_number}: "))

    while mark < 0 or mark > 100:
        print("Marks must be between 0 and 100.")
        mark = float(input(f"Enter marks for subject {subject_number}: "))

    total_marks += mark

average = total_marks / number_of_subjects

if average >= 90:
    grade = "A+"
elif average >= 80:
    grade = "A"
elif average >= 70:
    grade = "B"
elif average >= 60:
    grade = "C"
elif average >= 50:
    grade = "D"
else:
    grade = "F"

if average >= 50:
    result = "Pass"
else:
    result = "Fail"

print("\n=== Student Report ===")
print("Student Name:", student_name)
print("Total Marks:", total_marks)
print("Average Marks:", round(average, 2))
print("Grade:", grade)
print("Result:", result)