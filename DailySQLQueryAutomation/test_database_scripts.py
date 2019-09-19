import query, database_results
import failed_testcases_report as ft

def test_query_1():
    result = database_results.get_db_results(query.db_conn, query.query_1)
    actual_result = len(result)
    query_list = []
    query_list.append(result)
    err_msg = "Datacheck - query_1 failed with following errors"
    send_email(query_list, err_msg) if actual_result else ''
    assert actual_result == 0, "{0}\n {1}".format(err_msg, result) if actual_result else "Need to contact the Support team to investigate on the issue"
    
def test_query_2():
    result = database_results.get_db_results(query.db_conn, query.query_2)
    actual_result = len(result)
    query_list = []
    query_list.append(result)
    err_msg = "Datacheck - query_2 failed with following errors"
    send_email(query_list, err_msg) if actual_result else ''
    assert actual_result == 0, "{0}\n {1}".format(err_msg, result) if actual_result else "Need to contact the Support team to investigate on the issue"


def send_email(query_list, subject):
    parsed_data = ft.html_parser(query_list, subject)
    ft.send_email(parsed_data, subject)
