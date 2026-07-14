/*==============================================================*/
/* DBMS name:      Sybase SQL Anywhere 11                       */
/* Created on:     2023/5/26 12:57:17                           */
/*==============================================================*/


if exists(select 1 from sys.sysforeignkey where role='FK_BORROW/G_BORROW/GI_BOOKS') then
    alter table "borrow/give"
       delete foreign key "FK_BORROW/G_BORROW/GI_BOOKS"
end if;

if exists(select 1 from sys.sysforeignkey where role='FK_BORROW/G_BORROW/GI_USERS') then
    alter table "borrow/give"
       delete foreign key "FK_BORROW/G_BORROW/GI_USERS"
end if;

if exists(select 1 from sys.sysforeignkey where role='FK_INFORMAT_INFORMAT _MANAGEME') then
    alter table "informat manage"
       delete foreign key "FK_INFORMAT_INFORMAT _MANAGEME"
end if;

if exists(select 1 from sys.sysforeignkey where role='FK_INFORMAT_INFORMAT _USERS') then
    alter table "informat manage"
       delete foreign key "FK_INFORMAT_INFORMAT _USERS"
end if;

if exists(select 1 from sys.sysforeignkey where role='FK_MANAGEME_MANAGE_BOOKS') then
    alter table management
       delete foreign key FK_MANAGEME_MANAGE_BOOKS
end if;

if exists(
   select 1 from sys.systable 
   where table_name='Users'
     and table_type in ('BASE', 'GBL TEMP')
) then
    drop table Users
end if;

if exists(
   select 1 from sys.systable 
   where table_name='books'
     and table_type in ('BASE', 'GBL TEMP')
) then
    drop table books
end if;

if exists(
   select 1 from sys.systable 
   where table_name='borrow/give'
     and table_type in ('BASE', 'GBL TEMP')
) then
    drop table "borrow/give"
end if;

if exists(
   select 1 from sys.systable 
   where table_name='informat manage'
     and table_type in ('BASE', 'GBL TEMP')
) then
    drop table "informat manage"
end if;

if exists(
   select 1 from sys.systable 
   where table_name='management'
     and table_type in ('BASE', 'GBL TEMP')
) then
    drop table management
end if;

/*==============================================================*/
/* Table: Users                                                 */
/*==============================================================*/
create table Users 
(
   UsersNAME            char(20)                       null,
   UsersID              char(20)                       not null,
   UsersMIMA            char(30)                       null,
   constraint PK_USERS primary key (UsersID)
);

/*==============================================================*/
/* Table: books                                                 */
/*==============================================================*/
create table books 
(
   bookISBN             char(20)                       not null,
   bookNAME             char(20)                       null,
   bookPRICE            char(5)                       null,
   bookjieyueren        char(20)                       null,
   bookdata             char(13)                       null,
   constraint PK_BOOKS primary key (bookISBN)
);

drop table books


/*==============================================================*/
/* Table: "borrow/give"                                         */
/*==============================================================*/
create table "borrow/give" 
(
   UsersID              char(20)                       not null,
   bookISBN             char(20)                       not null,
   kucun                char(20)                       null,
   constraint "PK_BORROW/GIVE" primary key (UsersID, bookISBN)
);

/*==============================================================*/
/* Table: "informat manage"                                     */
/*==============================================================*/
create table "informat manage" 
(
   ManID                char(20)                       not null,
   UsersID              char(20)                       not null,
   constraint "PK_INFORMAT MANAGE" primary key (ManID, UsersID)
);

/*==============================================================*/
/* Table: management                                            */
/*==============================================================*/
create table management 
(
   ManName              char(20)                       null,
   ManID                char(20)                       not null,
   bookISBN             char(20)                       null,
   ManMIMA              char(30)                       null,
   constraint PK_MANAGEMENT primary key (ManID)
);

alter table "borrow/give"
   add constraint "FK_BORROW/G_BORROW/GI_BOOKS" foreign key (bookISBN)
      references books (bookISBN)
      on update restrict
      on delete restrict;

alter table "borrow/give"
   add constraint "FK_BORROW/G_BORROW/GI_USERS" foreign key (UsersID)
      references Users (UsersID)
      on update restrict
      on delete restrict;

alter table "informat manage"
   add constraint "FK_INFORMAT_INFORMAT _MANAGEME" foreign key (ManID)
      references management (ManID)
      on update restrict
      on delete restrict;

alter table "informat manage"
   add constraint "FK_INFORMAT_INFORMAT _USERS" foreign key (UsersID)
      references Users (UsersID)
      on update restrict
      on delete restrict;

alter table management
   add constraint FK_MANAGEME_MANAGE_BOOKS foreign key (bookISBN)
      references books (bookISBN)
      on update restrict
      on delete restrict;

