# ANP-PROJECT - [EN :uk:]

---

# Vinicius Guerra e Ribas -  Energy Sector Analyst
[Energy Engineer (UnB)](https://www.unb.br/) │ [Data Scientist and Analytics (USP)](https://www5.usp.br/)


## [:email: E-mail](mailto:viniciusgribas@gmail.com?Subject=%5BANP-PROJECT%5D%20-%20Contact)│ [:dart: Linkedin](https://www.linkedin.com/in/vinicius-guerra-e-ribas/) │[:octocat: GitHub](https://github.com/viniciusgribas) 

---

# [:computer: Project Notebook](https://github.com/viniciusgribas/ANP-PROJECT/blob/main/Codigos_Python/Notebook_Master.ipynb)

---

## PART 1 - INTRODUCTION
The programming languages used were PYTHON and VBA EXCEL.

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

### 1.2 PROJECT FLOW

1) At first, the files are only available in excel `".XLS"` format, under the name [`"vendas-combustíveis-m3.xls"`](https://github.com/viniciusgribas/ANP-PROJECT/tree/main/assets).

2) Within this initial file, there are two pivot tables that are the target. These are:

    - Pivot Table 1 ) "Vendas, pelas distribuidoras, dos derivados combustíveis de petróleo por Unidade da Federação e produto - 2000-2020 (m3)"
    
    - Pivot Table 2 ) "Vendas, pelas distribuidoras, de óleo diesel por tipo e Unidade da Federação - 2013-2020 (m3)"

3) This data, presented by the pivot tables, does not have its data source easily accessible in another spreadsheet. Also, the data is not available through the Excel shortcut: PivotTableTools>Analyze>Change Data Source. This shows the need to extract them using Excel's own `VBA` programming language. The advantage of extracting them this way is not only the reduced time for processes that could be long, but the possibility of applying them via `python`, through the *[`xlwings library`](https://www.xlwings.org/)*.

     - The worksheet, once opened, has by default only one sheet, called "plan1".
     
     - The macros created in VBA are available in the folder [`"\ANP-PROJECT\Codigos_VBA"`](https://github.com/viniciusgribas/ANP-PROJECT/tree/main/Codigos_VBA).
     - To extract this data, 4 macros were created in VBA. These are presented and described in **PART 2 EXCEL**

4) Once all the *VBA - MACROS* have been created, they can be called by `python` and applied there via *[`xlwings library`](https://www.xlwings.org/)*.

5) After applying the Macros on python, the end products of the extraction are two files in `"CSV-UTF8"`:

    - [`PlanConsolidada1.CSV`](https://github.com/viniciusgribas/ANP-PROJECT/tree/main/assets)
       
    - [`PlanConsolidada2.CSV`](https://github.com/viniciusgribas/ANP-PROJECT/tree/main/assets)

6) These files were managed via the *[`Pandas Library`](https://pandas.pydata.org/)* from python. Having the descriptive in **PART 3 PYTHON**

7) Finally, the final product of this project is two files in `"CSV-UTF8"` available in the folder [`"\ANP-PROJECT\Planilhas Finais"`](https://github.com/viniciusgribas/ANP-PROJECT/tree/main/Planilhas%20Finais), according to the following table:

| Column     | Type      |
|------------|-----------|
| year_month | date      |
| uf         | string    |
| product    | string    |
| unit       | string    |
| volume     | double    |
| created_at | timestamp |

   - [`Sales_Of_Diesel_By_UF_And_Type.CSV`](https://github.com/viniciusgribas/ANP-PROJECT/tree/main/Planilhas%20Finais)

   - [`Sales_Of_Oil_Derivative_Fuels_By_UF_And_Product.CSV`](https://github.com/viniciusgribas/ANP-PROJECT/tree/main/Planilhas%20Finais)


![image](https://user-images.githubusercontent.com/63691577/161165472-a0c20a8b-d68e-4d6b-a11c-c717bfbb0aa7.png)

