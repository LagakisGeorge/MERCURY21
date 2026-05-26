USE ANTOPD2020
SELECT G.ID_NUM AS DocID, R.TITLOS AS  DocDescr,R.SYNT_TITL AS DocCode ,G.ATIM AS DocNum,SUBSTRING(G.ATIM,2,7) AS DocAA,
G.HME AS DocDate,G.ATIM AS AlterDoc,PROOD_AJ AS LineNum,PELKOD AS CustCode,
P.EPO AS CustName,P.AFM AS  CustAFM,P.EPA AS CustProfession,P.DOY AS CustDOY,P.XRVMA AS CustZipCode,P.DIE AS CustAddress,P.POL AS CustCity,
1 AS VatStatus,LEFT(TRP,1) AS Payment ,TRP AS PaymentDescr ,KODE AS EidosCode,ONOMA AS EidosDescription,MONA AS UNIT,POSO AS Qty,TIMM AS Price,KAU_AJIA AS NetValue,
(CASE WHEN FPA=1 THEN 13 WHEN FPA=2 THEN 24 WHEN FPA=3 THEN 6 ELSE 0 END) AS FPAPrc,
MIK_AJIA-KAU_AJIA AS VATValue,G.EKPT AS DiscVal,0 AS DiscVal2 ,
G.HME AS InsertDate ,
GETDATE() AS LastUpdate,(CASE WHEN AKYROMENO=0 THEN 1 ELSE 0 END) AS IsActive,
NULL AS OrderID 
FROM EGGTIM G 
INNER JOIN PARASTAT R ON  LEFT(G.ATIM,1)=R.EIDOS
INNER JOIN TIM T ON T.ID_NUM=G.ID_NUM
INNER JOIN PEL P ON G.PELKOD=P.KOD AND G.EIDOS=P.EIDOS 
WHERE LEFT(KODE,2)='71'
/*
       
 

LastUpdate              à Είναι σημαντικό το συγκεκριμένο πεδίο να μην έχει μόνο τιμές στο σκέλος της ημερομηνίας αλλά και της ώρας. (DATETIME)

 

IsActive                   àΑφορά στο αν το παραστατικό είναι ‘Κανονικό’ / ‘Ακυρωμένο’ / ‘Ακυρωτικό’. Το πεδίο αυτό πρέπει να έχει τιμή 1 όταν πρόκειται για ‘Κανονικό’ και τιμή 0 όταν πρόκειται για ‘Ακυρωμένο’. Τα ‘Ακυρωτικά’ τα βγάζουμε εκτός.

    à Το προσθέτετε σαν πεδίο ως έχει


*/

