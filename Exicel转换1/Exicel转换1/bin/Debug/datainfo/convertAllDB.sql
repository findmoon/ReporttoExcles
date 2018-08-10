PRAGMA foreign_keys=OFF;
BEGIN TRANSACTION;
CREATE TABLE CPHTable(
    id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
    moduleid INTEGER NOT NULL,
    headid INTEGER NOT NULL,
    avgCPH INTEGER,
    prior_productionCPH INTEGER,
    high_precisionCPH INTEGER,
    special TEXT
);
INSERT INTO CPHTable VALUES(1,1,2,10000,'null','null','null');
INSERT INTO CPHTable VALUES(2,1,3,6000,'null','null','null');
INSERT INTO CPHTable VALUES(3,1,8,5500,'null','null','null');
INSERT INTO CPHTable VALUES(4,1,4,4200,'null','null','null');
INSERT INTO CPHTable VALUES(5,1,7,15000,'null','null','null');
INSERT INTO CPHTable VALUES(6,2,2,9400,'null','null','null');
INSERT INTO CPHTable VALUES(7,2,3,5600,'null','null','null');
INSERT INTO CPHTable VALUES(8,2,8,5500,'null','null','null');
INSERT INTO CPHTable VALUES(9,2,4,4200,'null','null','null');
INSERT INTO CPHTable VALUES(10,2,7,15000,'null','null','null');
INSERT INTO CPHTable VALUES(11,201,12,26000,'null','null','null');
INSERT INTO CPHTable VALUES(12,201,9,22500,'null','null','null');
INSERT INTO CPHTable VALUES(13,201,7,18000,'null','null','null');
INSERT INTO CPHTable VALUES(14,201,2,10500,'null','null','null');
INSERT INTO CPHTable VALUES(15,201,10,6800,'null','null','null');
INSERT INTO CPHTable VALUES(16,201,11,9500,'null','null','null');
INSERT INTO CPHTable VALUES(17,201,3,6800,'null','null','null');
INSERT INTO CPHTable VALUES(18,201,8,5500,'null','null','null');
INSERT INTO CPHTable VALUES(19,201,4,4200,'null','null','null');
INSERT INTO CPHTable VALUES(20,501,12,26000,'null','null','null');
INSERT INTO CPHTable VALUES(21,501,9,22500,'null','null','null');
INSERT INTO CPHTable VALUES(22,501,7,18000,'null','null','null');
INSERT INTO CPHTable VALUES(23,501,2,10500,'null','null','null');
INSERT INTO CPHTable VALUES(24,501,10,6800,'null','null','null');
INSERT INTO CPHTable VALUES(25,501,11,9500,'null','null','null');
INSERT INTO CPHTable VALUES(26,501,3,6800,'null','null','null');
INSERT INTO CPHTable VALUES(27,501,8,5500,'null','null','null');
INSERT INTO CPHTable VALUES(28,501,4,4200,'null','null','null');
INSERT INTO CPHTable VALUES(29,202,12,26000,'null','null','null');
INSERT INTO CPHTable VALUES(30,202,9,22500,'null','null','null');
INSERT INTO CPHTable VALUES(31,202,7,18000,'null','null','null');
INSERT INTO CPHTable VALUES(32,202,13,13000,'null','null','null');
INSERT INTO CPHTable VALUES(33,202,2,10500,'null','null','null');
INSERT INTO CPHTable VALUES(34,202,10,6800,'null','null','null');
INSERT INTO CPHTable VALUES(35,202,11,9500,'null','null','null');
INSERT INTO CPHTable VALUES(36,202,3,6800,'null','null','null');
INSERT INTO CPHTable VALUES(37,202,8,5500,'null','null','null');
INSERT INTO CPHTable VALUES(38,202,4,4200,'null','null','null');
INSERT INTO CPHTable VALUES(39,502,12,26000,'null','null','null');
INSERT INTO CPHTable VALUES(40,502,9,22500,'null','null','null');
INSERT INTO CPHTable VALUES(41,502,7,18000,'null','null','null');
INSERT INTO CPHTable VALUES(42,502,13,13000,'null','null','null');
INSERT INTO CPHTable VALUES(43,502,2,10500,'null','null','null');
INSERT INTO CPHTable VALUES(44,502,10,6800,'null','null','null');
INSERT INTO CPHTable VALUES(45,502,11,9500,'null','null','null');
INSERT INTO CPHTable VALUES(46,502,3,6800,'null','null','null');
INSERT INTO CPHTable VALUES(47,502,8,5500,'null','null','null');
INSERT INTO CPHTable VALUES(48,502,4,4200,'null','null','null');
INSERT INTO CPHTable VALUES(49,601,15,35000,42000,22000,'H24S');
INSERT INTO CPHTable VALUES(50,601,15,35000,37500,22000,'H24G');
INSERT INTO CPHTable VALUES(51,601,15,35000,37500,22000,'null');
INSERT INTO CPHTable VALUES(52,601,12,26000,'null','null','null');
INSERT INTO CPHTable VALUES(53,601,9,24500,'null','null','null');
INSERT INTO CPHTable VALUES(54,601,7,18000,'null','null','null');
INSERT INTO CPHTable VALUES(55,601,2,11500,'null','null','null');
INSERT INTO CPHTable VALUES(56,601,19,7500,'null','null','null');
INSERT INTO CPHTable VALUES(57,601,10,7500,'null','null','null');
INSERT INTO CPHTable VALUES(58,601,18,10500,'null','null','null');
INSERT INTO CPHTable VALUES(59,601,11,9500,'null','null','null');
INSERT INTO CPHTable VALUES(60,601,3,6500,'null','null','null');
INSERT INTO CPHTable VALUES(61,601,17,6700,7400,'null','null');
INSERT INTO CPHTable VALUES(62,601,8,5500,'null','null','null');
INSERT INTO CPHTable VALUES(63,601,4,4200,'null','null','null');
INSERT INTO CPHTable VALUES(64,701,15,35000,42000,22000,'H24S');
INSERT INTO CPHTable VALUES(65,701,15,35000,37500,22000,'H24G');
INSERT INTO CPHTable VALUES(66,701,15,35000,37500,22000,'null');
INSERT INTO CPHTable VALUES(67,701,12,26000,'null','null','null');
INSERT INTO CPHTable VALUES(68,701,9,24500,'null','null','null');
INSERT INTO CPHTable VALUES(69,701,7,18000,'null','null','null');
INSERT INTO CPHTable VALUES(70,701,2,11500,'null','null','null');
INSERT INTO CPHTable VALUES(71,701,19,7500,'null','null','null');
INSERT INTO CPHTable VALUES(72,701,10,7500,'null','null','null');
INSERT INTO CPHTable VALUES(73,701,18,10500,'null','null','null');
INSERT INTO CPHTable VALUES(74,701,11,9500,'null','null','null');
INSERT INTO CPHTable VALUES(75,701,3,6500,'null','null','null');
INSERT INTO CPHTable VALUES(76,701,17,6700,7400,'null','null');
INSERT INTO CPHTable VALUES(77,701,8,5500,'null','null','null');
INSERT INTO CPHTable VALUES(78,701,4,4200,'null','null','null');
INSERT INTO CPHTable VALUES(79,602,15,35000,42000,22000,'H24S');
INSERT INTO CPHTable VALUES(80,602,15,35000,37500,22000,'H24G');
INSERT INTO CPHTable VALUES(81,602,15,35000,37500,22000,'null');
INSERT INTO CPHTable VALUES(82,602,12,26000,'null','null','null');
INSERT INTO CPHTable VALUES(83,602,9,24500,'null','null','null');
INSERT INTO CPHTable VALUES(84,602,7,18000,'null','null','null');
INSERT INTO CPHTable VALUES(85,602,13,13000,14000,'null','null');
INSERT INTO CPHTable VALUES(86,602,2,11500,'null','null','null');
INSERT INTO CPHTable VALUES(87,602,19,7500,'null',6200,'null');
INSERT INTO CPHTable VALUES(88,602,10,7500,'null','null','null');
INSERT INTO CPHTable VALUES(89,602,18,10500,'null','null','null');
INSERT INTO CPHTable VALUES(90,602,11,9500,'null','null','null');
INSERT INTO CPHTable VALUES(91,602,3,6500,'null','null','null');
INSERT INTO CPHTable VALUES(92,602,17,6700,7400,'null','null');
INSERT INTO CPHTable VALUES(93,602,8,5500,'null','null','null');
INSERT INTO CPHTable VALUES(94,602,4,4200,'null','null','null');
INSERT INTO CPHTable VALUES(95,702,15,35000,42000,22000,'H24S');
INSERT INTO CPHTable VALUES(96,702,15,35000,37500,22000,'H24G');
INSERT INTO CPHTable VALUES(97,702,15,35000,37500,22000,'null');
INSERT INTO CPHTable VALUES(98,702,12,26000,'null','null','null');
INSERT INTO CPHTable VALUES(99,702,9,24500,'null','null','null');
INSERT INTO CPHTable VALUES(100,702,7,18000,'null','null','null');
INSERT INTO CPHTable VALUES(101,702,13,13000,14000,'null','null');
INSERT INTO CPHTable VALUES(102,702,2,11500,'null','null','null');
INSERT INTO CPHTable VALUES(103,702,19,7500,'null',6200,'null');
INSERT INTO CPHTable VALUES(104,702,10,7500,'null','null','null');
INSERT INTO CPHTable VALUES(105,702,18,10500,'null','null','null');
INSERT INTO CPHTable VALUES(106,702,11,9500,'null','null','null');
INSERT INTO CPHTable VALUES(107,702,3,6500,'null','null','null');
INSERT INTO CPHTable VALUES(108,702,17,6700,7400,'null','null');
INSERT INTO CPHTable VALUES(109,702,8,5500,'null','null','null');
INSERT INTO CPHTable VALUES(110,702,4,4200,'null','null','null');
INSERT INTO CPHTable VALUES(111,101,9,17000,'null','null','null');
INSERT INTO CPHTable VALUES(112,102,9,17500,'null','null','null');
INSERT INTO CPHTable VALUES(113,101,7,16500,'null','null','null');
INSERT INTO CPHTable VALUES(114,102,7,17000,'null','null','null');
INSERT INTO CPHTable VALUES(115,101,2,10000,'null','null','null');
INSERT INTO CPHTable VALUES(116,102,2,10000,'null','null','null');
INSERT INTO CPHTable VALUES(117,101,11,7800,'null','null','null');
INSERT INTO CPHTable VALUES(118,102,11,7800,'null','null','null');
INSERT INTO CPHTable VALUES(119,101,3,6000,'null','null','null');
INSERT INTO CPHTable VALUES(120,102,3,6000,'null','null','null');
INSERT INTO CPHTable VALUES(121,101,4,3500,'null','null','null');
INSERT INTO CPHTable VALUES(122,102,4,3500,'null','null','null');
INSERT INTO CPHTable VALUES(123,1,11,7600,'null','null','null');
INSERT INTO CPHTable VALUES(124,1,8,5500,'null','null','null');
INSERT INTO CPHTable VALUES(125,1,4,3500,'null','null','null');
INSERT INTO CPHTable VALUES(126,1,9,15500,'null','null','null');
INSERT INTO CPHTable VALUES(127,2,11,7600,'null','null','null');
INSERT INTO CPHTable VALUES(128,2,8,5500,'null','null','null');
INSERT INTO CPHTable VALUES(129,2,4,3200,'null','null','null');
INSERT INTO CPHTable VALUES(130,2,9,15500,'null','null','null');
CREATE TABLE ModuleTable(
    moduleid INTEGER PRIMARY KEY,
    module TEXT,
    moduletype TEXT
, trayinfostring TEXT);
INSERT INTO ModuleTable VALUES(1,'M3','M3',NULL);
INSERT INTO ModuleTable VALUES(2,'M6','M6',NULL);
INSERT INTO ModuleTable VALUES(101,'M3S','M3S',NULL);
INSERT INTO ModuleTable VALUES(102,'M6S','M6S',NULL);
INSERT INTO ModuleTable VALUES(201,'M3-2','M3II',NULL);
INSERT INTO ModuleTable VALUES(202,'M6-2','M6II','LTray-LTTray-LT2Tray-LTCTray-MTray');
INSERT INTO ModuleTable VALUES(302,'M6-2SP','M6IISP',NULL);
INSERT INTO ModuleTable VALUES(501,'M3-2c','M3IIc',NULL);
INSERT INTO ModuleTable VALUES(502,'M6-2c','M6IIc',NULL);
INSERT INTO ModuleTable VALUES(601,'M3-3','M3III',NULL);
INSERT INTO ModuleTable VALUES(602,'M6-3','M6III','LTray-LTTray-LT2Tray-LTCTray-MTray');
INSERT INTO ModuleTable VALUES(701,'M3-3c','M3IIIc',NULL);
INSERT INTO ModuleTable VALUES(702,'M6-3c','M6IIIc',NULL);
INSERT INTO ModuleTable VALUES(902,'M6-3L','M6IIIL',NULL);
CREATE TABLE HeadTable(
    headid INTEGER PRIMARY KEY,
    head TEXT
);
INSERT INTO HeadTable VALUES(2,'H08');
INSERT INTO HeadTable VALUES(3,'H04');
INSERT INTO HeadTable VALUES(4,'H01');
INSERT INTO HeadTable VALUES(6,'OF');
INSERT INTO HeadTable VALUES(7,'H12S');
INSERT INTO HeadTable VALUES(8,'H02');
INSERT INTO HeadTable VALUES(9,'H12HS');
INSERT INTO HeadTable VALUES(10,'G04');
INSERT INTO HeadTable VALUES(11,'H04S');
INSERT INTO HeadTable VALUES(12,'V12');
INSERT INTO HeadTable VALUES(13,'H08M');
INSERT INTO HeadTable VALUES(15,'H24S');
INSERT INTO HeadTable VALUES(17,'H02F');
INSERT INTO HeadTable VALUES(18,'H04SF');
INSERT INTO HeadTable VALUES(19,'G04F');
INSERT INTO HeadTable VALUES(21,'V12D');
INSERT INTO HeadTable VALUES(25,'DX');
INSERT INTO HeadTable VALUES(101,'GL');
INSERT INTO HeadTable VALUES(301,'IH1-15');
CREATE TABLE ModuleHeadCphStructTable
(
    TimeStamp INTEGER,
    Indx INTEGER,
    ModuleTrayTypeGroup	TEXT,
    HeadTypeGroup TEXT,
    CPHListGroup TEXT,
    FOREIGN KEY(TimeStamp) REFERENCES EvaluationReportTable(TimeStamp)
);
CREATE TABLE EvaluationReportTable
(
    TimeStamp INTEGER PRIMARY KEY,
    machineKind	TEXT,
    machineName	TEXT,
    Jobname TEXT,
    ToporBot TEXT,
    targetConveyor TEXT,
    productionMode TEXT,
    summaryDTXml TEXT,
    ModulePictureByte BLOB,
    layoutDTXml	TEXT,
    feedrNozzleDTXml TEXT,
    AirPowerNetPictureByte	BLOB,
    ModuleTypeStrings TEXT,
    BaseStasticsJson TEXT,
    moduleStatisticsJson TEXT,
    HeadTypeStrings TEXT,
    HeadTypeNozzleDictJson TEXT
);
CREATE TABLE UserTable
(
    User Text PRIMARY KEY,
    Pwd TEXt NOT NULL,
    Level INTEGER DEFAULT 4,
    NickName TEXT,
    ExpireTime TEXT
);
INSERT INTO UserTable VALUES('Administrator','E3AFED0047B08059D0FADA10F400C1E5',3,NULL,'2018-11-30');
INSERT INTO UserTable VALUES('asksz','CC04D2CF11C7C034829334CA9B8A7A22',3,NULL,'2018-11-30');
INSERT INTO UserTable VALUES('askA','37F1EF11567B7F24CA9B6586914B32F2',0,NULL,'2019-05-01');
DELETE FROM sqlite_sequence;
INSERT INTO sqlite_sequence VALUES('CPHTable',130);
CREATE TRIGGER expiredValue_Tri
AFTER INSERT ON UserTable
BEGIN
    UPDATE UserTable SET ExpireTime=date('now','start of day','+90 day')
    WHERE User=new.User;
END;
COMMIT;
