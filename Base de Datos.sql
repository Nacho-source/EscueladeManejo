drop database if exists Escuela_Manejo;
create database Escuela_Manejo;
use Escuela_Manejo;

create table persona(
idPersona int auto_increment,
nombre varchar(15),
apellido varchar (15),
direccion varchar (30),
tel varchar (10),
email varchar (30),
dni varchar(10),
constraint pk_persona primary key (idPersona)
)engine=innodb;

create table alumno(
idAlumno int auto_increment,
idPersona int,
constraint pk_alumno primary key (idAlumno),
constraint fk_alumno foreign key (idPersona) references persona (idPersona)
)engine=innodb;

create table instructor(
idInstructor int auto_increment,
idPersona int,
FIngreso date,
CursosEspeciales boolean,
constraint pk_instructor primary key (idInstructor),
constraint fk_instructor foreign key (idPersona) references persona (idPersona)
)engine=innodb;

create table vehiculo(
CodVehiculo int auto_increment,
matricula varchar(10),
marca varchar(15),
modelo varchar(15),
utilizable boolean,
FRevision date,
constraint pk_vehiculo primary key (CodVehiculo)
)engine=innodb;

create table caja(
idCaja int auto_increment,
FCaja date,
saldoinicial float,
saldofinal float,
MontoEntrada float,
MontoSalidaMes float,
constraint pk_caja primary key (idCaja)
)engine=innodb;

create table curso(
NCurso int,
nombre varchar (30),
precio float,
practico boolean,
especial boolean,
constraint pk_curso primary key (NCurso)
)engine=innodb;

create table horario(
NHorario int,
Dia varchar(10),
HoraEntrada varchar(10),
HoraSalida varchar(10),
constraint pk_horario primary key (NHorario)
)engine=innodb;

create table CursoFecha(
idCF int,
NCurso int,
NHorario int,
constraint pk_cursofecha primary key (idCF),
constraint fk_cursofecha foreign key (NCurso) references curso (NCurso),
constraint fk2_cursofecha foreign key (NHorario) references horario (NHorario)
)engine=innodb;

create table CursoInstructor(
idCF int,
idInstructor int,
FInicio date,
constraint pk_CursoInstructor primary key (idCF, idInstructor),
constraint fk_CursoInstructor foreign key (idCF) references CursoFecha (idCF),
constraint fk2_CursoInstructor foreign key (idInstructor) references instructor (idInstructor)
)engine=innodb;

create table teorico(
NTeorico int,
NCurso int,
Cupo int,
constraint pk_teorico primary key (NTeorico),
constraint fk_teorico foreign key (NCurso) references curso (NCurso)
)engine=innodb;

create table AlumnoCurso(
idAlumno int,
idCF int,
FIngreso date,
aprobado boolean,
constraint pk_AlumnoCurso primary key (idAlumno, idCF),
constraint fk_AlumnoCurso foreign key (idAlumno) references alumno (idAlumno),
constraint fk2_AlumnoCurso foreign key (idCF) references cursofecha (idCF)
)engine=innodb;

create table AlumnoVehiculo(
idAlumno int,
idCF int,
CodVehiculo int,
constraint pk_AlumnoVehiculo primary key (idAlumno, CodVehiculo),
constraint fk_AlumnoVehiculo foreign key (idAlumno) references alumno(idAlumno),
constraint fk2_AlumnoVehiculo foreign key (CodVehiculo) references vehiculo (CodVehiculo),
constraint fk3_AlumnoVehiculo foreign key (idCF) references CursoFecha(idCF)
)engine=innodb;

create table factura(
NFactura int auto_increment,
idAlumno int,
FPago date,
constraint pk_factura primary key (NFactura, idAlumno),
constraint fk_factura foreign key (idAlumno) references alumno (idAlumno)
)engine=innodb;

create table DetalleFactura(
NFactura int,
NCurso int,
Monto float,
Observaciones varchar (100),
constraint pk_DetalleFactura primary key (NFactura, NCurso),
constraint fk_DetalleFactura foreign key (NFactura) references factura (NFactura),
constraint fk2_DetalleFactura foreign key (NCurso) references curso (NCurso)
)engine=innodb;

create table PagoSueldo(
CodSueldo int auto_increment,
idCF int,
idInstructor int,
Fecha date,
Monto float,
constraint pk_sueldo primary key (CodSueldo),
constraint fk_sueldo foreign key (idInstructor) references instructor (idInstructor),
constraint fk2_sueldo foreign key (idCF) references CursoFecha (idCF)
)engine=innodb;

create table ListaEspera(
NEspera int auto_increment,
idPersona int,
NCurso int,
FIngreso date,
constraint pk_ListaEspera primary key (NEspera),
constraint fk_ListaEspera foreign key (idPersona) references persona (idPersona),
constraint fk2_ListaEspera foreign key (NCurso) references curso (NCurso)
)engine=innodb;

create table pago(
NPago int auto_increment,
NFactura int,
FormaPago varchar(15),
FPago date,
monto float,
constraint pk_pago primary key (NPago),
constraint fk_pago foreign key (NFactura) references factura (NFactura)
)engine=innodb;

create table RegistroAsistencia(
CodAsis int auto_increment,
idPersona int,
idCF int,
Fecha date,
HoraEntrada varchar (10),
constraint pk_RegistroAsistencia primary key (CodAsis),
constraint fk_RegistroAsistencia foreign key (idCF) references CursoFecha(idCF),
constraint fk2_RegistroAsistencia foreign key (idPersona) references persona (idPersona)
)engine=innodb;

create table justificativo(
NJustificativo int auto_increment,
CodAsis int,
FJustificada date,
constraint pk_justificativo primary key (NJustificativo),
constraint fk_justificativo foreign key (CodAsis) references RegistroAsistencia (CodAsis)
)engine=innodb;

create table HorarioInstructor(
NHorario int,
idInstructor int,
constraint pk_HorarioInstructor primary key (NHorario, idInstructor),
constraint fk_HorarioInstructor foreign key (NHorario) references horario (NHorario),
constraint fk2_HorarioInstructor foreign key (idInstructor) references instructor (idInstructor)
)engine=innodb;

create table evaluacion(
CodEval int auto_increment,
idAlumno int,
NCurso int,
nota float,
idInstructor int,
constraint pk_evaluacion primary key (CodEval),
constraint fk_evaluacion foreign key (idAlumno) references alumno (idAlumno),
constraint fk2_evaluacion foreign key (NCurso) references curso (NCurso)
)engine=innodb;

create table contraseña(
NContraseña int auto_increment,
contraseña long,
constraint pk_contraseña primary key (NContraseña)
)engine=innodb;

create table rotura(
CodRotura int auto_increment,
Fecha date,
CodVehiculo int,
idAlumno int,
Pagada boolean,
Monto float,
constraint pk_rotura primary key (CodRotura),
constraint fk1_rotura foreign key (CodVehiculo) references vehiculo (CodVehiculo),
constraint fk2_rotura foreign key (idAlumno) references alumno (idAlumno)
)engine=innodb;

create table PagoRotura(
CodRotura int,
FPago date,
constraint pk_PagoRotura primary key (CodRotura, FPago),
constraint fk_PagoRotura foreign key (CodRotura) references rotura (CodROtura)
)engine=innodb;