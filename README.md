# Protect
$oExcel.Range("C1:C10").Locked = False With $oWorkbook.ActiveSheet     ; set required options here     ;.AllowFormattingCells = True     ;.AllowInsertingColumns = False     ;.AllowDeletingColumns = False          .Protect("paswword") EndWith
