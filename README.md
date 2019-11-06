# Interop Issues

## 1: Can't use Excel on .NET Core

### Repo steps

1. Edit NetCoreInteropIssue.csproj and set `TargetFramework` to `netcoreapp3.0`
2. `dotnet run`

### Expected behavior

Excel opens

### Actual behavior

App crashes:

```
Unhandled exception. Microsoft.CSharp.RuntimeBinder.RuntimeBinderException: 'System.__ComObject' does not contain a definition for 'Range'
   at CallSite.Target(Closure , CallSite , Object , String )
   at System.Dynamic.UpdateDelegates.UpdateAndExecute2[T0,T1,TRet](CallSite site, T0 arg0, T1 arg1)
   at Microsoft.Csv.ExcelExtensions.ImportCsvDocument(Application a, CsvDocument csvDocument) in NetCoreInteropIssue\CSV\ExcelExtensions.cs:line 74
   at Microsoft.Csv.ExcelExtensions.LoadCsvDocument(Application a, CsvDocument csvDocument) in NetCoreInteropIssue\CSV\ExcelExtensions.cs:line 54
   at Microsoft.Csv.ExcelExtensions.ViewInExcel(CsvDocument csvDocument) in NetCoreInteropIssue\CSV\ExcelExtensions.cs:line 26
   at NetCoreInteropIssue.Program.Main(String[] args) in NetCoreInteropIssue\Program.cs:line 10
```

## 2: Not embedding PIA will fail to compile

### Repo steps

1. Under **Dependencies | COM** right click the
   **Interop.Microsoft.Office.Interop.Excel**  and select **Properties**
2. Set **Embed Interop Types** to **False**
3. Recompile

### Expected behavior

The code compiles but now references the PIA instead of embedding.

### Actual behavior

You get these compiler errors in `ExcelExtensions.cs`

```
Line 91   CS1061: 'object' does not contain a definition for 'Range' and no accessible extension method 'Range' accepting a first argument of type 'object' could be found (are you missing a using directive or an assembly reference?)

Line 92   CS1061: 'object' does not contain a definition for 'Range' and no accessible extension method 'Range' accepting a first argument of type 'object' could be found (are you missing a using directive or an assembly reference?)

Line 116  CS1061: 'object' does not contain a definition for 'ListObjects' and no accessible extension method 'ListObjects' accepting a first argument of type 'object' could be found (are you missing a using directive or an assembly reference?)

Line 113  CS1061: 'object' does not contain a definition for 'End' and no accessible extension method 'End' accepting a first argument of type 'object' could be found (are you missing a using directive or an assembly reference?)

Line 114  CS1061: 'object' does not contain a definition for 'End' and no accessible extension method 'End' accepting a first argument of type 'object' could be found (are you missing a using directive or an assembly reference?)
```