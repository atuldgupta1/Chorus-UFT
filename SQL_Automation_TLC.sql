-- Connect Primary Location
--Input : Central Office
select NAOWN.NADBGE_2.tlc from NAOWN.NADBGE_2@PRODDATANA,LocationDB.Address_Search@ALM as1
where
not exists (select null from PRDINV.product_instance pi
    where SUBSTR(pi.LOCATION_REF_A,1,1)='1' and PI.LOCATION_REF_A = NAOWN.NADBGE_2.TLC)
and 
NOT EXISTS (select NULL from OM.order_product_instance opi
                join OM.product_order po on po.id =  opi.product_order_id
                join OM.customer_order_item_location coil on coil.customer_order_item_id = po.customer_order_item_id
                join OM.customer_order_item coi on coil.customer_order_item_id = coi.id
                join OM.Customer_order co on co.id = coi.Customer_order_id
                WHERE coil.location_ref = to_char( NAOWN.NADBGE_2.TLC)
                AND co.customer_order_state_substt_id not in (32,34,35,36,37,2,3,4,5,6,7,8,9,10,11,12))
AND AVAILABILITY_STATUS = 'Ready'
AND INSTALL_TYPE like 'Standard Install%%'
AND CONSENT_REQUIRED = 'N'
AND DESIGN_REQUIRED = 'N'
AND BUILD_REQUIRED = 'N'
AND ZONE_TYPE in ('UFB','Non-UFB')
and central_office = 'TPO'
And as1.TLC = NAOWN.NADBGE_2.tlc
AND ROWNUM <= 1


--Connect Additional ONT
--Input : Central_Office
select NAOWN.NADBGE_2.tlc from NAOWN.NADBGE_2@PRODDATANA
where NAOWN.NADBGE_2.tlc in
(select   pi.location_ref_A from prdinv.product_instance pi,Prdinv.product_instance_chr pic,prdcat.product_offer po
where pi.end_date is null
and pi.id = pic.product_instance_id
and pic.value like 'ONT%'
and po.id = pi.product_offer_id
and po.offer_type = 'Primary'
--No Inflight on that Instance
and exists  (select * from om.order_product_instance opi 
                  join OM.product_order po on opi.product_order_id = po.id
                  join om.customer_order_item coi on po.customer_order_item_id = coi.id
                  join om.customer_order co on coi.customer_order_id = co.id
                  where opi.ext_product_instance_id_ref = pi.external_ref
                  And co.customer_order_state_substt_id not in (0,1,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,36,37)) -- Inflight state
-- Move Primary Solution
and NOT EXISTS  (select null from OM.product_order po
                  join om.customer_order_item coi on po.customer_order_item_id = coi.id
                  join om.customer_order co on coi.customer_order_id = co.id
                  where pi.id = po.product_instance_id
                  And co.customer_order_state_substt_id not in (32,34,35,36,37,2,3,4,5,6,7,8,9,10,11,12)) -- Non Inflight State
-- No Inflight On that location
and Not Exists (select * from om.customer_order co, om.customer_order_item coi, om.customer_order_item_location coil
                where co.id = coi.customer_order_id
                and coi.id = coil.customer_order_item_id
                and coil.location_ref = to_char(pi.location_ref_A)
                and co.customer_order_state_substt_id in (0,1,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,36,37)) -- Inflight State
group by pi.location_ref_A
having count(*) = 1) 
AND AVAILABILITY_STATUS = 'Ready'
AND CONSENT_REQUIRED = 'N'
AND DESIGN_REQUIRED = 'N'
AND BUILD_REQUIRED = 'N'
AND ZONE_TYPE in ('UFB','Non-UFB')
and central_office = 'HSN'
AND ROWNUM <= 1

--Location for Connect Secondary ie no existing secondary
--Input : Central Office
select NAOWN.NADBGE_2.tlc from NAOWN.NADBGE_2@PRODDATANA 
where NAOWN.NADBGE_2.tlc in
(select  distinct pi.location_ref_A from prdinv.product_instance pi,Prdinv.product_instance_chr pic,prdcat.product_offer po
where pi.end_date is null
and pi.id = pic.product_instance_id
and pic.value like 'ONT%'
and po.id = pi.product_offer_id
and po.offer_type = 'Primary'
--and po.id = 27  --(In any particualy primary product is required)
--and pi.customer_ref = :CustomerReference -- Vod =- 54457865 Snap = 54530114 (if the primary instance needs to belong to particular RSP)
and exists  (select * from om.order_product_instance opi 
                  join OM.product_order po on opi.product_order_id = po.id
                  join om.customer_order_item coi on po.customer_order_item_id = coi.id
                  join om.customer_order co on coi.customer_order_id = co.id
                  where opi.ext_product_instance_id_ref = pi.external_ref
                  And co.customer_order_state_substt_id not in (0,1,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,36,37)) -- Inflight state
and not exists(select  distinct pi1.location_ref_A from prdinv.product_instance pi1,product_offer po1
				where pi1.end_date is null
				and po1.id = pi1.product_offer_id
				and po1.offer_type = 'Secondary'
				and pi1.location_ref_A = pi.location_ref_A)
and NOT EXISTS  (select null from OM.product_order po
                  join om.customer_order_item coi on po.customer_order_item_id = coi.id
                  join om.customer_order co on coi.customer_order_id = co.id
                  where pi.id = po.product_instance_id
                  And co.customer_order_state_substt_id not in (32,34,35,36,37,2,3,4,5,6,7,8,9,10,11,12)) -- Non Inflight State
and Not Exists (select * from om.customer_order co, om.customer_order_item coi, om.customer_order_item_location coil
                where co.id = coi.customer_order_id
                and coi.id = coil.customer_order_item_id
                and coil.location_ref = to_char(pi.location_ref_A)
                and co.customer_order_state_substt_id in (0,1,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,36,37)))
AND AVAILABILITY_STATUS = 'Ready'
AND CONSENT_REQUIRED = 'N'
AND DESIGN_REQUIRED = 'N'
AND BUILD_REQUIRED = 'N'
AND ZONE_TYPE in ('UFB','Non-UFB')
and central_office = 'HSN'
AND ROWNUM <= 1


-- Location for Modify Attribute/Change Offer/Transfer/Disconnect Primary for RSP
--Input : Central Office And Customer Reference
select NAOWN.NADBGE_2.tlc from NAOWN.NADBGE_2@PRODDATANA 
where NAOWN.NADBGE_2.tlc in
(select  distinct pi.location_ref_A from prdinv.product_instance pi,Prdinv.product_instance_chr pic,prdcat.product_offer po
where pi.end_date is null
and pi.id = pic.product_instance_id
and pic.value like 'ONT%'
and po.id = pi.product_offer_id
and po.offer_type = 'Primary'
--and po.id = 27
and pi.customer_ref = 54765460  -- Vod =- 54457865 Snap = 54530114 Airnet = 54765460
and exists  (select * from om.order_product_instance opi 
                  join OM.product_order po on opi.product_order_id = po.id
                  join om.customer_order_item coi on po.customer_order_item_id = coi.id
                  join om.customer_order co on coi.customer_order_id = co.id
                  where opi.ext_product_instance_id_ref = pi.external_ref
                  And co.customer_order_state_substt_id not in (0,1,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,36,37)) -- Inflight state
and NOT EXISTS  (select null from OM.product_order po
                  join om.customer_order_item coi on po.customer_order_item_id = coi.id
                  join om.customer_order co on coi.customer_order_id = co.id
                  where pi.id = po.product_instance_id
                  And co.customer_order_state_substt_id not in (32,34,35,36,37,2,3,4,5,6,7,8,9,10,11,12)) -- Non Inflight State
and Not Exists (select * from om.customer_order co, om.customer_order_item coi, om.customer_order_item_location coil
                where co.id = coi.customer_order_id
                and coi.id = coil.customer_order_item_id
                and coil.location_ref = to_char(pi.location_ref_A)
                and co.customer_order_state_substt_id in (0,1,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,36,37)))
AND AVAILABILITY_STATUS = 'Ready'
AND CONSENT_REQUIRED = 'N'
AND DESIGN_REQUIRED = 'N'
AND BUILD_REQUIRED = 'N'
AND ZONE_TYPE in ('UFB','Non-UFB')
and central_office = 'HSN'
AND ROWNUM <= 1


--Location with existing primary and Secondary Instance
--Input : Central Office
select NAOWN.NADBGE_2.tlc from NAOWN.NADBGE_2@PRODDATANA 
where NAOWN.NADBGE_2.tlc in
(select  distinct pi.location_ref_A from prdinv.product_instance pi,Prdinv.product_instance_chr pic,prdcat.product_offer po
where pi.end_date is null
and pi.id = pic.product_instance_id
and pic.value like 'ONT%'
and po.id = pi.product_offer_id
and po.offer_type = 'Primary'
--and po.id = 27
--and pi.customer_ref = :CustomerReference -- Vod =- 54457865 Snap = 54530114
and exists  (select * from om.order_product_instance opi 
                  join OM.product_order po on opi.product_order_id = po.id
                  join om.customer_order_item coi on po.customer_order_item_id = coi.id
                  join om.customer_order co on coi.customer_order_id = co.id
                  where opi.ext_product_instance_id_ref = pi.external_ref
                  And co.customer_order_state_substt_id not in (0,1,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,36,37)) -- Inflight state
and Exists (select  distinct pi1.location_ref_A from prdinv.product_instance pi1,product_offer po1
				where pi1.end_date is null
        and pi1.customer_ref = :CustomerReference
				and po1.id = pi1.product_offer_id
				and po1.offer_type = 'Secondary'
				and pi1.location_ref_A = pi.location_ref_A)
and NOT EXISTS  (select null from OM.product_order po
                  join om.customer_order_item coi on po.customer_order_item_id = coi.id
                  join om.customer_order co on coi.customer_order_id = co.id
                  where pi.id = po.product_instance_id
                  And co.customer_order_state_substt_id not in (32,34,35,36,37,2,3,4,5,6,7,8,9,10,11,12)) -- Non Inflight State
and Not Exists (select * from om.customer_order co, om.customer_order_item coi, om.customer_order_item_location coil
                where co.id = coi.customer_order_id
                and coi.id = coil.customer_order_item_id
                and coil.location_ref = to_char(pi.location_ref_A)
                and co.customer_order_state_substt_id in (0,1,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,36,37)))
AND AVAILABILITY_STATUS = 'Ready'
AND CONSENT_REQUIRED = 'N'
AND DESIGN_REQUIRED = 'N'
AND BUILD_REQUIRED = 'N'
AND ZONE_TYPE in ('UFB','Non-UFB')
and central_office = 'HSN'
AND ROWNUM <= 1


-- Connect Primary Intact ONT Location
-- Input : Central Office
select NAOWN.NADBGE_2.tlc from NAOWN.NADBGE_2@PRODDATANA 
where NAOWN.NADBGE_2.tlc in  
(select  distinct pi.location_ref_A from prdinv.product_instance pi,Prdinv.product_instance_chr pic
where pi.end_date is not null
and pi.id = pic.product_instance_id
and pic.value like 'ONT%' 
and not exists (Select * from prdinv.product_instance pi1 where pi.location_ref_a = pi1.location_ref_A and pi1.End_Date is null)
and not exists (select NULL from OM.order_product_instance opi
                join OM.product_order po on po.id =  opi.product_order_id
                join OM.customer_order_item_location coil on coil.customer_order_item_id = po.customer_order_item_id
                join OM.customer_order_item coi on coil.customer_order_item_id = coi.id
                join OM.Customer_order co on co.id = coi.Customer_order_id
                WHERE coil.location_ref = pi.location_ref_A
                AND co.customer_order_state_substt_id not in (32,34,35,36,37,2,3,4,5,6,7,8,9,10,11,12)))
AND AVAILABILITY_STATUS = 'Ready'
AND CONSENT_REQUIRED = 'N'
AND DESIGN_REQUIRED = 'N'
AND BUILD_REQUIRED = 'N'
AND ZONE_TYPE in ('UFB','Non-UFB')
and central_office = 'HSN'
AND ROWNUM <= 1