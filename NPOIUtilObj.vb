Imports Microsoft.VisualBasic
Imports NPOI.XSSF.UserModel
Imports System.Data
Imports System.Data.SqlClient
Imports System.Reflection
Imports System.IO
Imports System.Web

Public Class NPOIUtilObj
    Private m_SheetName As String = ""
    Private m_ImgPath As String = ""
    Private m_SkipRow As Integer = 3

    ''' <summary>
    ''' 唯寫，傳入Excel Sheet的名稱
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public WriteOnly Property SheetName As String
        Set(value As String)
            m_SheetName = value
        End Set
    End Property

    ''' <summary>
    ''' 唯寫，傳入插入圖檔的實體路徑
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public WriteOnly Property ImgPath As String
        Set(value As String)
            m_ImgPath = value
        End Set
    End Property

    ''' <summary>
    ''' 唯寫，傳入上方要空幾筆空間
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public WriteOnly Property SkipRow As Integer
        Set(value As Integer)
            m_SkipRow = value
        End Set
    End Property

    ''' <summary>
    ''' 依據傳入的DataTable，產生Excel的WorkBook，並傳回
    ''' </summary>
    ''' <param name="Dt">DataTable</param>
    ''' <returns>成功傳回Excel WorkBook</returns>
    ''' <remarks></remarks>
    Public Function DtToWorkBook(ByRef Dt As DataTable) As XSSFWorkbook
        Dim book As New XSSFWorkbook
        Dim sheet As XSSFSheet = book.CreateSheet(m_SheetName)

        If Dt.Rows.Count > 0 Then
            If m_ImgPath <> "" Then
                InertImgtoExcel(book, sheet)
            End If

            Dim x As Integer = 0
            Dim y As Integer = 0

            Dim rw As XSSFRow = sheet.CreateRow(m_SkipRow)

            '建制head
            For Each col As DataColumn In Dt.Columns
                rw.CreateCell(x).SetCellValue(col.ColumnName)
                x += 1
            Next

            y = m_SkipRow + 1
            Dim xsrw As XSSFRow
            For Each rwDt As DataRow In Dt.Rows
                xsrw = sheet.CreateRow(y)

                For x = 0 To Dt.Columns.Count - 1
                    xsrw.CreateCell(x).SetCellValue(rwDt.Item(x).ToString)
                Next

                y += 1
            Next

        Else
            Throw New Exception("DataTable無資料")
        End If
        Return book

    End Function

    Public Function DtAddToWorkBook(ByRef book As XSSFWorkbook, ByRef Dt As DataTable) As String
        Try
            Dim sheet As XSSFSheet = book.CreateSheet(m_SheetName)

            If Dt.Rows.Count > 0 Then
                If m_ImgPath <> "" Then
                    InertImgtoExcel(book, sheet)
                End If

                Dim x As Integer = 0
                Dim y As Integer = 0

                Dim rw As XSSFRow = sheet.CreateRow(m_SkipRow)

                '建制head
                For Each col As DataColumn In Dt.Columns
                    rw.CreateCell(x).SetCellValue(col.ColumnName)
                    x += 1
                Next

                y = m_SkipRow + 1
                Dim xsrw As XSSFRow
                For Each rwDt As DataRow In Dt.Rows
                    xsrw = sheet.CreateRow(y)

                    For x = 0 To Dt.Columns.Count - 1
                        xsrw.CreateCell(x).SetCellValue(rwDt.Item(x).ToString)
                    Next

                    y += 1
                Next

            Else
                Throw New Exception("DataTable無資料")
            End If

            Return "Success"

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' 依據傳入的DataReader，產生Excel的WorkBook並傳回
    ''' </summary>
    ''' <param name="Dr">傳入的DataReader</param>
    ''' <returns>成功回傳Excel的WorkBook</returns>
    ''' <remarks></remarks>
    Public Function DrToWorkBook(ByRef Dr As SqlDataReader) As XSSFWorkbook
        Dim book As New XSSFWorkbook
        Try
            If Dr.HasRows Then
                Dim sheet As XSSFSheet = book.CreateSheet(m_SheetName)

                If m_ImgPath <> "" Then
                    InertImgtoExcel(book, sheet)
                End If

                Dim x As Integer = 0
                Dim y As Integer = 0

                Dim rw As XSSFRow = sheet.CreateRow(3)

                '建制head

                For x = 0 To Dr.FieldCount - 1
                    rw.CreateCell(x).SetCellValue(Dr.GetName(x))
                Next

                y = 4
                Dim xsrw As XSSFRow
                While Dr.Read()
                    xsrw = sheet.CreateRow(y)
                    For x = 0 To Dr.FieldCount - 1
                        xsrw.CreateCell(x).SetCellValue(Dr.Item(x).ToString)
                    Next
                    y += 1
                End While

            End If
            Dr.Close()

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return book
    End Function

    ''' <summary>
    ''' 依據傳入的物件(集合)，轉換成Excel的WorkBook並傳回
    ''' </summary>
    ''' <param name="oObjs">傳入的物件集合</param>
    ''' <returns>成功回傳Excel的WorkBook</returns>
    ''' <remarks></remarks>
    Public Function ObjToWorkBook(ByVal oObjs As IEnumerable(Of Object)) As XSSFWorkbook

        Try
            Dim book As New XSSFWorkbook
            Dim sheet As XSSFSheet = book.CreateSheet(m_SheetName)

            If oObjs.Count > 0 Then
                If m_ImgPath <> "" Then
                    InertImgtoExcel(book, sheet)
                End If
                Dim x As Integer = 0
                Dim y As Integer = 0

                Dim rw As XSSFRow = sheet.CreateRow(3)

                '建制head
                For Each pty As PropertyInfo In oObjs(0).GetType().GetProperties()
                    rw.CreateCell(x).SetCellValue(pty.Name)
                    x += 1
                Next

                y = 4
                Dim xsrw As XSSFRow
                For Each o As Object In oObjs
                    xsrw = sheet.CreateRow(y)
                    x = 0
                    For Each pty As PropertyInfo In o.GetType().GetProperties()
                        xsrw.CreateCell(x).SetCellValue(pty.GetValue(o).ToString)
                        x += 1
                    Next
                    'For x = 0 To o.GetType.GetProperties.Count - 1
                    '    xsrw.CreateCell(x).SetCellValue(o.)
                    'Next
                    y += 1
                Next
            Else
                Throw New Exception("物件無資料")
            End If

            Return book
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' 圖檔放入Excel
    ''' </summary>
    ''' <param name="book">XSSFWorkbook</param>
    ''' <param name="sheet">XSSFSheet</param>
    ''' <remarks></remarks>
    Private Sub InertImgtoExcel(ByRef book As XSSFWorkbook, ByRef sheet As XSSFSheet)
        Dim bytes As Byte() = System.IO.File.ReadAllBytes(m_ImgPath)
        Dim pictureIdx As Integer = book.AddPicture(bytes, XSSFWorkbook.PICTURE_TYPE_GIF)

        Dim drawing As XSSFDrawing = sheet.CreateDrawingPatriarch()
        Dim helper As XSSFCreationHelper = book.GetCreationHelper
        Dim anchor As XSSFClientAnchor
        '設定圖片位置
        'anchor = helper.CreateClientAnchor()
        anchor = New XSSFClientAnchor(dx1:=5, dy1:=2, dx2:=0, dy2:=0, col1:=0, row1:=0, col2:=0, row2:=0)

        Dim pict As XSSFPicture = drawing.CreatePicture(anchor, pictureIdx)
        pict.Resize()
    End Sub

    Public Sub SaveWorkBook(ByRef Response As HttpResponse, ByVal book As XSSFWorkbook, ByVal FileName As String)
        Dim ms As New MemoryStream
        book.Write(ms)
        Response.AddHeader("Content-Disposition", String.Format("attachment; filename=" & FileName & ".xlsx"))
        Response.BinaryWrite(ms.ToArray())
        book = Nothing
        ms.Close()
        ms.Dispose()
        Response.End()
    End Sub


End Class
