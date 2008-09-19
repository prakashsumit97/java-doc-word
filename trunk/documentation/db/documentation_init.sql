drop table compound if exists;
drop table member if exists;

create table compound (
	id varchar primary key ,
	kind varchar,
	prot varchar,
	name varchar,
	detaileddescription varchar,
	basecompounddef varchar
);

create table member (
	id varchar primary key ,
	compound_id varchar,
	prot varchar,
	kind varchar,
	static_prefix varchar,
	name varchar,
	type_name varchar,
	detaileddescription varchar,
	params varchar
);

