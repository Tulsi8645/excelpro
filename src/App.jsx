import { useState } from "react";

const excelScript = `
  Public Starting_Range_AD As Date ' Starting date in AD
Public BS As Variant ' Declaring BS as a Variant array
Public Array_Size As Integer ' Array size variable

Sub BS_Database()
    ' The corresponding AD start date for 2070-01-01 BS is 2013-04-14 AD
    Starting_Range_AD = DateValue("2013-04-14")

    ' Initializing the BS database from 2070 BS to 2090 BS
    BS = Array( _
        Array(2070, 31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 30, 30), _
        Array(2071, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30), _
        Array(2072, 31, 32, 31, 32, 31, 30, 30, 29, 30, 29, 30, 30), _
        Array(2073, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31), _
        Array(2074, 31, 31, 31, 32, 31, 31, 30, 29, 30, 29, 30, 30), _
        Array(2075, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30), _
        Array(2076, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 30), _
        Array(2077, 31, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31), _
        Array(2078, 31, 31, 31, 32, 31, 31, 30, 29, 30, 29, 30, 30), _
        Array(2079, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30), _
        Array(2080, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 30), _
        Array(2081, 31, 31, 32, 32, 31, 30, 30, 30, 29, 30, 29, 31), _
        Array(2082, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 30, 30), _
        Array(2083, 31, 31, 32, 31, 31, 30, 30, 30, 29, 30, 30, 30), _
        Array(2084, 31, 31, 32, 31, 31, 30, 30, 30, 29, 30, 30, 30), _
        Array(2085, 31, 32, 31, 32, 30, 31, 30, 30, 29, 30, 30, 30), _
        Array(2086, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 30, 30), _
        Array(2087, 31, 31, 32, 31, 31, 31, 30, 30, 29, 30, 30, 30), _
        Array(2088, 30, 31, 32, 32, 30, 31, 30, 30, 29, 30, 30, 30), _
        Array(2089, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 30, 30), _
        Array(2090, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 30, 30) _
    )

    Array_Size = UBound(BS) ' Store array size
End Sub

Function ParseDateString(InputDate As String) As Variant
    Dim DateParts As Variant

    If InStr(InputDate, ".") > 0 Then
        DateParts = Split(InputDate, ".")
        If UBound(DateParts) = 2 Then
            ParseDateString = DateParts
            Exit Function
        End If
    End If

    ' Return an error message if no valid format is found
    ParseDateString = Array("Error")
End Function

Function AD2BS(AD_Input As String) As String
    Dim AD As Date
    Dim DD_Gap As Double, DD_total As Double
    Dim i As Integer, ii As Integer
    Dim DateParts As Variant

    Call BS_Database ' Load BS data

    ' Parse AD input
    DateParts = ParseDateString(AD_Input)
    If DateParts(0) = "Error" Then
        AD2BS = "Error! Invalid AD date format..."
        Exit Function
    End If

    ' Convert to date
    On Error Resume Next
    AD = DateSerial(Val(DateParts(0)), Val(DateParts(1)), Val(DateParts(2)))
    If Err.Number <> 0 Then
        AD2BS = "Error! Invalid AD date..."
        Exit Function
    End If
    On Error GoTo 0

    DD_Gap = DateDiff("d", Starting_Range_AD, AD) ' Calculate difference in days

    ' If the AD date is out of range, return an error
    If DD_Gap < 0 Or DD_Gap > 5113 Then
        AD2BS = "Error! Out of range..."
        Exit Function
    End If

    DD_total = 0

    For i = 0 To Array_Size
        For ii = 1 To 12
            DD_total = DD_total + BS(i)(ii)

            If DD_total > DD_Gap Then
                AD2BS = BS(i)(0) & "." & Format(ii, "00") & "." & Format((DD_Gap - (DD_total - BS(i)(ii)) + 1), "00")
                Exit Function
            End If
        Next ii
    Next i
End Function

Function BS2AD(BS_Input As String) As String
    Dim i As Integer, ii As Integer
    Dim Total_Days As Double
    Dim BS_Year As Integer, BS_Month As Integer, BS_Day As Integer
    Dim DateParts As Variant

    Call BS_Database ' Load BS data

    ' Parse BS input
    DateParts = ParseDateString(BS_Input)
    If DateParts(0) = "Error" Then
        BS2AD = "Error! Invalid BS date format..."
        Exit Function
    End If

    ' Convert the parts into year, month, and day
    BS_Year = Val(DateParts(0))
    BS_Month = Val(DateParts(1))
    BS_Day = Val(DateParts(2))

    Total_Days = 0 ' Total days counter

    ' Find the year in the database
    For i = 0 To Array_Size
        If BS(i)(0) = BS_Year Then ' Matching the input year
            Exit For
        End If
        ' Add all days of previous years
        For ii = 1 To 12
            Total_Days = Total_Days + BS(i)(ii)
        Next ii
    Next i

    ' Validate if the year was found
    If i > UBound(BS) Then
        BS2AD = "Error! Year out of range..."
        Exit Function
    End If

    ' Validate month
    If BS_Month < 1 Or BS_Month > 12 Then
        BS2AD = "Error! Invalid month..."
        Exit Function
    End If

    ' Add days for the months before the input month
    For ii = 1 To BS_Month - 1
        Total_Days = Total_Days + BS(i)(ii)
    Next ii

    ' Validate day
    If BS_Day < 1 Or BS_Day > BS(i)(BS_Month) Then
        BS2AD = "Error! Invalid day..."
        Exit Function
    End If

    ' Add the day offset
    Total_Days = Total_Days + BS_Day - 1

    ' Compute the AD date by adding Total_Days to the starting date (2013-04-14)
    BS2AD = Format(DateAdd("d", Total_Days, Starting_Range_AD), "yyyy.mm.dd")
End Function
`;

export default function App() {
  const [copied, setCopied] = useState(false);

  const handleCopy = () => {
    navigator.clipboard.writeText(excelScript).then(() => {
      setCopied(true);
      setTimeout(() => setCopied(false), 2000); // Reset after 2s
    });
  };

  return (
    <div style={styles.container}>
      <div style={styles.card}>
        <h1 style={styles.title}>Excel Date Converter</h1>
        <button onClick={handleCopy} style={styles.button}>
          {copied ? "Copied!" : "Copy"}
        </button>
        <pre style={styles.codeBlock}>{excelScript}</pre>
      </div>
    </div>
  );
}

const styles = {
  container: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    height: "auto",
    backgroundColor: "#f4f4f4",
    padding: "20px",
  },
  card: {
    position: "relative",
    backgroundColor: "white",
    padding: "20px",
    borderRadius: "10px",
    boxShadow: "0px 4px 6px rgba(0, 0, 0, 0.1)",
    width: "100%",
    maxWidth: "800px",
  },
  title: {
    fontSize: "24px",
    fontWeight: "bold",
    marginBottom: "10px",
  },
  button: {
    position: "absolute",
    top: "20px",
    right: "20px",
    backgroundColor: "#007BFF",
    color: "white",
    padding: "8px 16px",
    border: "none",
    borderRadius: "5px",
    cursor: "pointer",
    transition: "background 0.3s",
  },
  codeBlock: {
    marginTop: "10px",
    backgroundColor: "#eaeaea",
    padding: "10px",
    borderRadius: "5px",
    fontSize: "14px",
    overflowX: "auto",
    whiteSpace: "pre-wrap",
  },
};
