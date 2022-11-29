import statistic
import graphs


input_user = statistic.InputConnect()
vacancies_array = statistic.DataSet.csv_reader(input_user.file_name)
if len(statistic.vacancies_array) == 0:
    statistic.do_exit('Ничего не найдено')

data = statistic.DataDictionaries()
data.update_data(vacancies_array, input_user.profession)
data.print()

report = statistic.Report(data)
report.generate_excel()

data = graphs.DataDictionaries()
data.update_data(vacancies_array, input_user.profession)
data.print()

report = graphs.Report(data)
report.generate_image()