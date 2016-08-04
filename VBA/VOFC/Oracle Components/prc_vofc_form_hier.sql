drop table vofc_form_hierarchy;
create table vofc_form_hierarchy 
 (seq number,
  section_id number,
  question_id number,
  relation_id number,
  parent_id number,
  nlevel number,
  relation_type varchar2(4000),
  html_type varchar2(4000),
  object_type varchar2(4000),
  vulnerability_group number,
  relation_name varchar2(4000),
  tooltip varchar2(4000));

declare
 vFmHier  vofc_form_hierarchy%rowtype;
 vSeq number := 0;
 vNum number;
 isGroup boolean;
 vGrpOptionID number;
 vGroupID number;
 vGrpParentID number;
 
 procedure insert_hier is
 begin
  vSeq := vSeq + 1;
  vFmHier.seq := vSeq;
  insert into  vofc_form_hierarchy
    values vFmHier;
 end insert_hier;

begin

for x in (select  level,
        RELATION_ID,                
        PARENT_ID,                  
        RELATION_ORDER,             
        RELATION_TYPE,
        HTML_TYPE,
        VULNERABILITY_GROUP,
        VUL_GROUP_PRIMARY,
        RELATION_NAME,
        TOOLTIP,
       -- DECODE(RELATION_NAME,null,null,LPAD(' ',4*level) || RELATION_NAME) as RELATION_NAME_HUMAN,
        SYS_CONNECT_BY_PATH(relation_id, '_') as id_TRAIL
   from    psa_assessment.VOFC_RELATIONS
   start with  relation_id in (select distinct q.sectionid 
   from psa_assessment.formquestions q, psa_assessment.formoptions o, psa_assessment.vofc_relations v 
   where q.questionid = o.questionid and o.optionid = nvl(v.parent_id,0) and v.relation_type = 'VULNERABILITY')
   connect by prior RELATION_ID =PARENT_ID
   order siblings by RELATION_ORDER, relation_ID) loop
 if x.relation_type = 'SECTION' then
  vFmHier := null;
  vFmHIer.relation_Type := x.relation_type;
  vFmHIER.relation_name := x.relation_name;
  vFmHier.section_id := x.relation_Id;
  vFmHier.nlevel := x.level;
  insert_hier;
 elsif x.relation_type = 'QUESTION' then
  isGroup := false;
  vFmHier.Object_type := null;
  vFmHIer.relation_Type := x.relation_type;
  vFmHier.relation_name := x.relation_name;
  vFmHier.question_Id := x.relation_id;
  vFmHier.nlevel := x.level;
  vFmHier.relation_id := null;
  vFmHier.parent_id := null;
  vFmHier.HTML_TYPE := x.html_type;
  vFmHier.VULNERABILITY_GROUP := null;
  vFmHIer.TOOLTIP := null;
  insert_hier;
 elsif x.relation_type = 'VULNERABILITY' and x.vulnerability_group is null then
  isGroup := false;
  vGroupID := null;
  vFmHier.relation_name := x.relation_name;
  vFmHIer.relation_Type := x.relation_type;
  vFmHier.nlevel := x.level;
  vFmHier.relation_id := x.relation_id;
  vFmHier.parent_id := x.parent_id;
  select html_type into vFMHier.object_type from psa_assessment.vofc_relations
    where relation_id = x.parent_id;
  vFmHier.HTML_TYPE := x.html_type;
  vFmHier.VULNERABILITY_GROUP := null;
  vFmHIer.TOOLTIP := null;
  insert_hier;
 elsif x.relation_type = 'VULNERABILITY-GROUP' then
  isGroup := true;
  vFmHier.relation_name := x.relation_name;
  vFmHIer.relation_Type := x.relation_type;
  vFmHier.nlevel := x.level;
  vFmHier.relation_id := x.relation_id;
  vFmHier.parent_id := x.parent_id;
  select html_type into vFMHier.object_type from psa_assessment.vofc_relations
    where relation_id = x.parent_id;
  vFmHier.HTML_TYPE := x.html_type;
  vGroupID := x.vulnerability_group;
  vFmHier.VULNERABILITY_GROUP := x.vulnerability_group;
  vFmHIer.TOOLTIP := null;
  insert_hier;
 elsif x.relation_type = 'OPTION' and isGroup then
  vGrpParentID := x.parent_id;
 elsif x.relation_type = 'VULNERABILITY' and x.vulnerability_group is not null then
  vFmHier.relation_name := x.relation_name;
  vFmHIer.relation_Type := 'VULNERABILITY GRP MEMBER';
  vFmHier.nlevel := x.level;
  select html_type into vFMHier.object_type from psa_assessment.vofc_relations
    where relation_id = x.parent_id;
  vFmHier.relation_id := x.relation_id;
  vFmHier.parent_id := vGrpParentID;
  vFmHier.HTML_TYPE := x.html_type;
  vFmHier.VULNERABILITY_GROUP := x.vulnerability_group;
  vFmHIer.TOOLTIP := null;
  insert_hier;
 elsif x.relation_type in ('SUB','OFC') then
   vFmHier.object_type := null;
   vFmHier.nlevel := x.level;
   vFmHier.relation_type := x.relation_type;
   vFmHier.relation_id := x.relation_Id;
   vFmHier.parent_id := x.parent_id;
   vFmHier.HTML_TYPE := x.html_type;
   vFmHier.VULNERABILITY_GROUP := vGroupID;
   vFmHIer.RELATION_NAME := x.relation_name;
   vFmHIer.TOOLTIP := x.tooltip;
   insert_hier;
 else
  select count(*) into vNum from vofc_form_hierarchy
   where x.parent_id in (relation_id);
  if vNum > 0 then
   vFmHier.object_type := null;
   vFmHier.nlevel := x.level;
   vFmHier.relation_type := x.relation_type;
   vFmHier.relation_id := x.relation_Id;
   vFmHier.parent_id := x.parent_id;
   vFmHier.HTML_TYPE := x.html_type;
   vFmHier.VULNERABILITY_GROUP := vGroupID;
   vFmHIer.RELATION_NAME := x.relation_name;
   vFmHIer.TOOLTIP := x.tooltip;
   insert_hier;
  end if;
-- elsif x.relation_type = 'OPTION' and isGroup then
--  vGrpOptionID := x.parent_id;
 end if;
end loop;
end;
/
commit;