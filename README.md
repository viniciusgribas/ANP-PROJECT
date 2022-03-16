# ANP-PROJECT

# Vinicius Guerra e Ribas -  Energy Sector Analyst
[Energy Engineer (UnB)](https://www.unb.br/)│ [Data Scientist and Analytics (USP)](https://www5.usp.br/)



## [e-mail](viniciusgribas@gmail.com)│ [Linkedin](https://www.linkedin.com/in/vinicius-guerra-e-ribas/) │[GitHub](https://github.com/viniciusgribas) │ [Porfolio - (in progress)](https://viniciusgribas.github.io/portfolio/)



## PART 1 - INTRODUCTION
The programming languages used were PYTHON and VBA EXCEL.

This project is available in this [link.](https://github.com/raizen-analytics/data-engineering-test/blob/master/TEST.md) It was first proposed for a test.

Seeking to simplify and make clear the flow of activities to obtain the final product, this notebook has been divided into 4 parts.

> PART 1 - INTRODUCTION
 -  Containing a short summary of how the project was developed, the basics and bibliography.

> PART 2 - EXCEL VBA
 - Introducing the `VBA` formulas developed in excel to be called in `python`.

> PART 3 - PYTHON
 - Presenting the `python` code used for this project.

> PART 4 - CONCLUSION
 - Final considerations of the project.

### 1.1 AUXILIARY BIBLIOGRAPHY
 - https://www.automateexcel.com/vba-code-examples/
 - https://www.rondebruin.nl/index.htm
 - https://www.xlwings.org/
 - https://pandas.pydata.org/docs/
 - http://timgolden.me.uk/pywin32-docs/contents.html
 - https://github.com/wesm/pydata-book
 - https://github.com/fzumstein/python-for-excel
 - https://github.com/raizen-analytics/data-engineering-test

### 1.2 PROJECT FLOW

1) At first, the files are only available in excel `".XLS"` format, under the name "vendas-combustíveis-m3.xls".
2) Within this initial file, there are two pivot tables that are the target. These are:

    - Pivot Table 1 ) "Vendas, pelas distribuidoras, dos derivados combustíveis de petróleo por Unidade da Federação e produto - 2000-2020 (m3)"
    
    - Pivot Table 2 ) "Vendas, pelas distribuidoras, de óleo diesel por tipo e Unidade da Federação - 2013-2020 (m3)"

3) This data, presented by the pivot tables, does not have its data source easily accessible in another spreadsheet. Also, the data is not available through the Excel shortcut: PivotTableTools>Analyze>Change Data Source. This shows the need to extract them using Excel's own `VBA` programming language. The advantage of extracting them this way is not only the reduced time for processes that could be long, but the possibility of applying them via `python`, through the `"wlwings" library`.

- The worksheet, once opened, has by default only one sheet, called "plan1".
- The macros created in VBA are available in the folder `"\ANP\Codigos_VBA"`.
-  To extract this data, 4 macros were created in VBA. These are presented and described in **PART 2 EXCEL**

4) Once all the *VBA - MACROS* have been created, they can be called by `python` and applied there via `xlwings library`.

5) After applying the Macros on python, the end products of the extraction are two files in `"CSV-UTF8"`:

 - `PlanConsolidada1.CSV`

 - `PlanConsolidada2.CSV`

6) These files were managed via the *[Pandas Library](https://pandas.pydata.org/)* from python. Having the descriptive in **PART 3 PYTHON**

7) Finally, the final product of this project is two files in `"CSV-UTF8"` available in the folder `"\ANP\Planilhas Finais"`, according to the [project's initial proposal](https://github.com/raizen-analytics/data-engineering-test/blob/master/TEST.md):

 - `Sales_Of_Diesel_By_UF_And_Type.CSV`

 - `Sales_Of_Oil_Derivative_Fuels_By_UF_And_Product.CSV`
