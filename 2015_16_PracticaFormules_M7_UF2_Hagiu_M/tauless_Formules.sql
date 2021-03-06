create table Formules(
	codi_formula varchar(255) not null,
	nom_formula varchar(255),
	totalPes_grams integer,
	data_creacio DATE,
	activaOno varchar(1),		
	primary key(codi_formula)
	
)

create table Elements(
	codi_element varchar(255) not null,
	nom_element varchar(255),
	data_creacio DATE,
	primary key(codi_element)
)


create table Composicio(
	codi_formula varchar(255) not null,
	codi_element varchar(255) not null,
	quantitat_grams integer
);

INSERT INTO Formules (codi_formula, nom_formula, totalPes_grams, data_creacio, activaOno)VALUES ('A000', 'Acido sulfurico', 35, GETDATE(), 1);
INSERT INTO Elements (codi_element, nom_element, data_creacio)VALUES ('aa000', 'Hidrogeno', GETDATE());
INSERT INTO Elements (codi_element, nom_element, data_creacio)VALUES ('bb111', 'Azufre', GETDATE());
INSERT INTO Elements (codi_element, nom_element, data_creacio)VALUES ('cc222', 'Oxigeno', GETDATE());
INSERT INTO Composicio (codi_formula, codi_element, quantitat_grams)VALUES ('A000', 'aa000', 10);
INSERT INTO Composicio (codi_formula, codi_element, quantitat_grams)VALUES ('A000', 'bb111', 5);
INSERT INTO Composicio (codi_formula, codi_element, quantitat_grams)VALUES ('A000', 'cc222', 20);

INSERT INTO Formules (codi_formula, nom_formula, totalPes_grams, data_creacio, activaOno)VALUES ('B111', 'Acido cloroso', 20, GETDATE(), 0);
INSERT INTO Elements (codi_element, nom_element, data_creacio)VALUES ('dd333', 'Cloro', GETDATE());
INSERT INTO Composicio (codi_formula, codi_element, quantitat_grams)VALUES ('B111', 'aa000', 5);
INSERT INTO Composicio (codi_formula, codi_element, quantitat_grams)VALUES ('B111', 'dd333', 5);
INSERT INTO Composicio (codi_formula, codi_element, quantitat_grams)VALUES ('B111', 'cc222', 10);