using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JiraApplication {
    using Atlassian.Jira;
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;

    class JiraTest {
        public JiraTest() {
            Jira jiraConn = new Jira("https://issues.tolling.us/", "aporrello", "3Francis");
            Debug.Print(jiraConn.GetIssue("PANYNJ-760").Assignee);

            //IEnumerable<Atlassian.Jira.Issue> jiraIssues = jiraConn.GetIssuesFromJql(jqlString, 999);
            //
            //foreach (var issue in jiraIssues) {
            //   Debug.WriteLine(issue.Key.Value + " -- " + issue.Summary);
            //}

        }

        static string PrepareJqlbyDates(string beginDate, string endDate) {
            string jqlString = "project = PRJ AND status = Closed AND resolved >= " + beginDate + " AND resolved <= " + endDate;
            return jqlString;
        }
    }

}
