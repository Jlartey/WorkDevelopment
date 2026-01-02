INSERT INTO "NavigationLink" ("NavigationLinkID", "NavigationLinkName", "ParentNavLinkID", "MenuElementTypeID", "PosInLink", "MenuProp", "LinkDetail", "LinkUrl", "NavigationLinkCatID", "NavigationLinkTypeID", "UserAccessibleID", "LinkOpenStyle", "OtherInfo", "TitleDesc", "KeyPrefix") VALUES ('HIM001-N011T-01-39', 'GHS DHIMS Report', 'HIM001-N011T-01', '06', 3, 'Vert||Left||Over||RePosition||3C', 'GHS DHIMS Report', '&', 'N001', 'AdminViews', 'USA001', 'IN||NAV||800||600||Close', '-', '-', NULL);


INSERT INTO "NavigationLink" ("NavigationLinkID", "NavigationLinkName", "ParentNavLinkID", "MenuElementTypeID", "PosInLink", "MenuProp", "LinkDetail", "LinkUrl", "NavigationLinkCatID", "NavigationLinkTypeID", "UserAccessibleID", "LinkOpenStyle", "OtherInfo", "TitleDesc", "KeyPrefix") VALUES ('HIM001-N011T-01-40', 'All DHIMS Reports', 'HIM001-N011T-01-39', '13', 1, 'Vert||Left||Click||RePosition||3C', 'All DHIMS Reports', 'wpgPrtPrintInputFilter.asp?PrintLayoutName=DHIMSReport&PositionForTableName=WorkingDay&WorkingDayID=', 'N001', 'Preferences1', 'USA001', 'IN||NAV||800||600||Close', '-', '-', NULL);
INSERT INTO "NavigationLink" ("NavigationLinkID", "NavigationLinkName", "ParentNavLinkID", "MenuElementTypeID", "PosInLink", "MenuProp", "LinkDetail", "LinkUrl", "NavigationLinkCatID", "NavigationLinkTypeID", "UserAccessibleID", "LinkOpenStyle", "OtherInfo", "TitleDesc", "KeyPrefix") VALUES ('HIM001-N011T-01-41', 'DHIMS Report [New]', 'HIM001-N011T-01-39', '13', 2, 'Vert||Left||Click||RePosition||3C', 'DHIMS Report [New]', 'wpgPrtPrintInputFilter.asp?PrintLayoutName=DHIMSReport2&PositionForTableName=WorkingDay&WorkingDayID=', 'N003', 'PresSchedDuration', 'USA001', 'IN||NAV||800||600||Close', '-', '-', NULL);

--DELETIONS MADE FROM APPOINTMENT CAT TABLE
INSERT INTO "AppointmentCat" ("AppointmentCatID", "AppointmentCatName", "Description", "KeyPrefix") VALUES ('A004', 'ULTRASOUND', '', '');
INSERT INTO "AppointmentCat" ("AppointmentCatID", "AppointmentCatName", "Description", "KeyPrefix") VALUES ('A005', 'Radiology', '', '');
INSERT INTO "AppointmentCat" ("AppointmentCatID", "AppointmentCatName", "Description", "KeyPrefix") VALUES ('A008', 'DENTAL', '', '');
INSERT INTO "AppointmentCat" ("AppointmentCatID", "AppointmentCatName", "Description", "KeyPrefix") VALUES ('A009', 'surgical', '', '');
INSERT INTO "AppointmentCat" ("AppointmentCatID", "AppointmentCatName", "Description", "KeyPrefix") VALUES ('A010', 'OPD', '', '');
INSERT INTO "AppointmentCat" ("AppointmentCatID", "AppointmentCatName", "Description", "KeyPrefix") VALUES ('A011', 'ARES', '', '');
INSERT INTO "AppointmentCat" ("AppointmentCatID", "AppointmentCatName", "Description", "KeyPrefix") VALUES ('A012', 'Orthopedics', 'OPD', '');
INSERT INTO "AppointmentCat" ("AppointmentCatID", "AppointmentCatName", "Description", "KeyPrefix") VALUES ('A013', 'EYE', '', '');
INSERT INTO "AppointmentCat" ("AppointmentCatID", "AppointmentCatName", "Description", "KeyPrefix") VALUES ('A014', 'EYE', '', '');
INSERT INTO "AppointmentCat" ("AppointmentCatID", "AppointmentCatName", "Description", "KeyPrefix") VALUES ('A015', 'OBS THEATRE', '', '');
INSERT INTO "AppointmentCat" ("AppointmentCatID", "AppointmentCatName", "Description", "KeyPrefix") VALUES ('A016', 'NEURO/ORTHOPAEDICS THEATRE', 'FOR CRANIOTOMY TOMORROW', '');
INSERT INTO "AppointmentCat" ("AppointmentCatID", "AppointmentCatName", "Description", "KeyPrefix") VALUES ('A017', 'GENERAL THEATRE', '', '');

INSERT INTO "AppointmentCatType" ("AppointmentCatTypeID", "AppointmentCatTypeName", "AppointmentCatID", "Description", "KeyPrefix") VALUES ('A001', 'OTHER CONSULTATION', 'A001', '', '');
INSERT INTO "AppointmentCatType" ("AppointmentCatTypeID", "AppointmentCatTypeName", "AppointmentCatID", "Description", "KeyPrefix") VALUES ('A014', 'PHYSICIAN SPECIALIST CLINIC', 'A001', '', '');
INSERT INTO "AppointmentCatType" ("AppointmentCatTypeID", "AppointmentCatTypeName", "AppointmentCatID", "Description", "KeyPrefix") VALUES ('A015', 'Family Physician Clinic', 'A001', '', '');
INSERT INTO "AppointmentCatType" ("AppointmentCatTypeID", "AppointmentCatTypeName", "AppointmentCatID", "Description", "KeyPrefix") VALUES ('A016', 'Renal Physician Clinic', 'A001', '', '');
INSERT INTO "AppointmentCatType" ("AppointmentCatTypeID", "AppointmentCatTypeName", "AppointmentCatID", "Description", "KeyPrefix") VALUES ('A017', 'Diabetic Clinic', 'A001', '', '');
INSERT INTO "AppointmentCatType" ("AppointmentCatTypeID", "AppointmentCatTypeName", "AppointmentCatID", "Description", "KeyPrefix") VALUES ('A018', 'Cardiology Clinic', 'A001', '', '');
INSERT INTO "AppointmentCatType" ("AppointmentCatTypeID", "AppointmentCatTypeName", "AppointmentCatID", "Description", "KeyPrefix") VALUES ('A019', 'Endocrine Clinic', 'A001', '', '');
INSERT INTO "AppointmentCatType" ("AppointmentCatTypeID", "AppointmentCatTypeName", "AppointmentCatID", "Description", "KeyPrefix") VALUES ('A020', 'Four For Men', 'A001', '', '');
INSERT INTO "AppointmentCatType" ("AppointmentCatTypeID", "AppointmentCatTypeName", "AppointmentCatID", "Description", "KeyPrefix") VALUES ('A021', 'Well Woman', 'A001', '', '');
INSERT INTO "AppointmentCatType" ("AppointmentCatTypeID", "AppointmentCatTypeName", "AppointmentCatID", "Description", "KeyPrefix") VALUES ('A022', 'Dermatology Clinic', 'A001', '', '');
INSERT INTO "AppointmentCatType" ("AppointmentCatTypeID", "AppointmentCatTypeName", "AppointmentCatID", "Description", "KeyPrefix") VALUES ('A023', 'Clinical Psychology', 'A001', '', '');
INSERT INTO "AppointmentCatType" ("AppointmentCatTypeID", "AppointmentCatTypeName", "AppointmentCatID", "Description", "KeyPrefix") VALUES ('A025', 'NEUROSURGERY', 'A001', '', '');
INSERT INTO "AppointmentCatType" ("AppointmentCatTypeID", "AppointmentCatTypeName", "AppointmentCatID", "Description", "KeyPrefix") VALUES ('A026', 'PHYSICIAN SPECIALIST CLINIC', 'A001', '', '');
INSERT INTO "AppointmentCatType" ("AppointmentCatTypeID", "AppointmentCatTypeName", "AppointmentCatID", "Description", "KeyPrefix") VALUES ('A027', 'Obstetrics / ANC', 'A001', '', '');
INSERT INTO "AppointmentCatType" ("AppointmentCatTypeID", "AppointmentCatTypeName", "AppointmentCatID", "Description", "KeyPrefix") VALUES ('A028', 'Radiology / Scan', 'A001', '', '');
INSERT INTO "AppointmentCatType" ("AppointmentCatTypeID", "AppointmentCatTypeName", "AppointmentCatID", "Description", "KeyPrefix") VALUES ('A029', 'Well Baby', 'A001', '', '');
INSERT INTO "AppointmentCatType" ("AppointmentCatTypeID", "AppointmentCatTypeName", "AppointmentCatID", "Description", "KeyPrefix") VALUES ('A030', 'Gyanaecology', 'A001', '', '');
INSERT INTO "AppointmentCatType" ("AppointmentCatTypeID", "AppointmentCatTypeName", "AppointmentCatID", "Description", "KeyPrefix") VALUES ('A031', 'Anaesthesia', 'A001', '', '');
INSERT INTO "AppointmentCatType" ("AppointmentCatTypeID", "AppointmentCatTypeName", "AppointmentCatID", "Description", "KeyPrefix") VALUES ('A032', 'PAEDIATRIC SURGERY', 'A001', '', '');
INSERT INTO "AppointmentCatType" ("AppointmentCatTypeID", "AppointmentCatTypeName", "AppointmentCatID", "Description", "KeyPrefix") VALUES ('A035', 'Nephrology', 'A001', NULL, NULL);
INSERT INTO "AppointmentCatType" ("AppointmentCatTypeID", "AppointmentCatTypeName", "AppointmentCatID", "Description", "KeyPrefix") VALUES ('A036', 'Haematology', 'A001', NULL, NULL);
INSERT INTO "AppointmentCatType" ("AppointmentCatTypeID", "AppointmentCatTypeName", "AppointmentCatID", "Description", "KeyPrefix") VALUES ('A037', 'Dietetics', 'A001', NULL, NULL);
INSERT INTO "AppointmentCatType" ("AppointmentCatTypeID", "AppointmentCatTypeName", "AppointmentCatID", "Description", "KeyPrefix") VALUES ('A038', 'Asthma Clinic', 'A001', NULL, NULL);
INSERT INTO "AppointmentCatType" ("AppointmentCatTypeID", "AppointmentCatTypeName", "AppointmentCatID", "Description", "KeyPrefix") VALUES ('A039', 'Chiropathy', 'A001', NULL, NULL);
INSERT INTO "AppointmentCatType" ("AppointmentCatTypeID", "AppointmentCatTypeName", "AppointmentCatID", "Description", "KeyPrefix") VALUES ('A040', 'Physiotherapy', 'A001', NULL, NULL);
INSERT INTO "AppointmentCatType" ("AppointmentCatTypeID", "AppointmentCatTypeName", "AppointmentCatID", "Description", "KeyPrefix") VALUES ('A041', 'ENT', 'A001', NULL, NULL);
INSERT INTO "AppointmentCatType" ("AppointmentCatTypeID", "AppointmentCatTypeName", "AppointmentCatID", "Description", "KeyPrefix") VALUES ('A042', 'Dental', 'A001', NULL, NULL);
INSERT INTO "AppointmentCatType" ("AppointmentCatTypeID", "AppointmentCatTypeName", "AppointmentCatID", "Description", "KeyPrefix") VALUES ('A043', 'Urology', 'A001', NULL, NULL);
INSERT INTO "AppointmentCatType" ("AppointmentCatTypeID", "AppointmentCatTypeName", "AppointmentCatID", "Description", "KeyPrefix") VALUES ('A044', 'Gastroenterology', 'A001', '', '');
INSERT INTO "AppointmentCatType" ("AppointmentCatTypeID", "AppointmentCatTypeName", "AppointmentCatID", "Description", "KeyPrefix") VALUES ('A045', 'Ophthalmology', 'A001', NULL, NULL);
INSERT INTO "AppointmentCatType" ("AppointmentCatTypeID", "AppointmentCatTypeName", "AppointmentCatID", "Description", "KeyPrefix") VALUES ('A046', 'Optometry', 'A001', NULL, NULL);
INSERT INTO "AppointmentCatType" ("AppointmentCatTypeID", "AppointmentCatTypeName", "AppointmentCatID", "Description", "KeyPrefix") VALUES ('A047', 'Paediatrics', 'A001', NULL, NULL);
INSERT INTO "AppointmentCatType" ("AppointmentCatTypeID", "AppointmentCatTypeName", "AppointmentCatID", "Description", "KeyPrefix") VALUES ('A048', 'Endoscopy', 'A001', '', '');
INSERT INTO "AppointmentCatType" ("AppointmentCatTypeID", "AppointmentCatTypeName", "AppointmentCatID", "Description", "KeyPrefix") VALUES ('A053', 'GENERAL SURGERY', 'A001', '', '');

INSERT INTO "AppointmentCatType" ("AppointmentCatTypeID", "AppointmentCatTypeName", "AppointmentCatID", "Description", "KeyPrefix") VALUES ('A003', 'EMERGENCY SURGERY', 'A003', '', '');
