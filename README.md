# BQL-api-for-Python
Contains the BQL function callable directly from Python. It needs the Excel add-in for Bloomberg

Refer to the BQL.query guide from Bloomberg. In this function the clause are called as follow:
	- let() is still let
	- get() is still get
	- for() is here universe
	- with() is here settings

when passing the clauses values just pass the string you would write inside the clause.

Example:

	BQL.Query(get("IS_EPS(FPR='2015A').VALUE") for("members(['SPX Index'])"))
	--> BQL(get = "IS_EPS(FPR='2015A').VALUE"), universe = "members(['SPX Index'])")