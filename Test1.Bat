SET cmddir="E:\Project\Lowis7\Lowis_Reports_Testing"
e:
cd %cmddir%
SET filename="E:\Project\Lowis7\Lowis_Reports_Testing\Lowis_Reports_Testing\bin\Debug\Lowis_Reports_Testing.dll"
rem mstest /testcontainer:%filename% /test:"VerifyBeamChartsLink"
rem  E:\Project\Lowis7\Lowis_Reports_Testing\Lowis_Reports_Testing\bin\Debug\Lowis_Reports_Testing.dll
mstest /testcontainer:%filename% /test:"espwellgrpconfig"
::mstest /testcontainer:%filename% /test:"VerifySeparataorLink"
::mstest /testcontainer:%filename% /test:"ImportWellmodelFiles"