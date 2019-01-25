Attribute VB_Name = "Module1"
Sub Easy():
 

    Dim Ticker As String

    Dim Total_Stock_Volume As Double

    Dim Summary_Table_Row As Integer



    Total_Stock_Volume = 0



    ' Keep track of the location for each row/line in the summary table

    Summary_Table_Row = 2

    

        Range("I1").Value = "Ticker"

        Range("J1").Value = "Total Stock Volume"

    

    ' Determine the Last Row

    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    

        For i = 2 To LastRow



            ' Check if we are still within the same value, if it is not...

            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then



                ' Set the Ticker Value

                Ticker = Cells(i, 1).Value



                ' Add to Total_Stock_Volume

                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value



                ' Print Ticker in the Summary Table

                Range("I" & Summary_Table_Row).Value = Ticker



                ' Print the Total_Stock_Volume to the Summary Table

                Range("J" & Summary_Table_Row).Value = Total_Stock_Volume



                ' Add one to the summary table row

                Summary_Table_Row = Summary_Table_Row + 1

                

                ' Reset Total_Stock_Volume

                Total_Stock_Volume = 0



            Else



                ' Add to the Total Stock Volume

                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value



            End If

              

        Next i
        
        End Sub
        

