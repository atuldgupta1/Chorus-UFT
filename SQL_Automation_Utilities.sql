-- To Find Available HO at the location
-- Inputs : Customer Reference Number and Location_Ref
select * from PRDINV.product_instance
where product_offer_id = 1
and location_ref_A = 'MDR'
and End_Date is null
and customer_ref = 54765460

-- To find loosing RSP for Transfer/Connect And Replace Scenario
-- Input : TLC
select cm.customer_name from cmown.customer_manager cm,prdinv.product_instance pi
where pi.External_ref = 'input'
And cm.customer_ref = pi.customer_ref

-- Query to Find Instance at the OLD Location for Move Primary Scenarios
-- Input : Old Location and Customer_Rel
select pi.external_ref from prdinv.product_instance pi 
where pi.location_ref_a = '100554063'
and End_Date is null
and customer_ref = 54765460


--Query to Find Asscoiated Disconnect Order for give Replace Order Type
SELECT co.ext_Customer_Order_Id_Ref from om.customer_order_item coi,om.customer_Order co
WHERE coi.product_instance_id =
(SELECT coi1.Replace_Product_instance_id from om.customer_order_item coi1,om.customer_order co1
WHERE co1.id = coi1.customer_order_id
AND coi1.Order_type like '%Replace%'
AND co1.ext_customer_order_id_ref = '100036099')
AND co.customer_order_state_substt_id in (0,1,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,36,37)
And Co.id = coi.customer_order_id


--Query to Find External Product ID for the Given Order Nos
Select pi.External_ref from om.customer_order co,om.customer_order_item coi,prdinv.product_instance pi
where co.ext_customer_order_id_ref = 100036100
and co.id = coi.customer_order_id
and coi.product_instance_id = pi.id
and pi.End_Date is null

-- Query to Find Available CSE PRoduct Offer based on selection of RSP/PRoduct OFfer/AIM
select * from prdcat.product_offer where
id in
(select cpo1.product_offer_id from prdcat.customer_product_offer cpo1
where cpo1.id in 
(Select por.CHILD_CUST_PROD_OFFER_ID from prdcat.product_offer_Relationship por
where por.PARENT_CUST_PROD_OFFER_ID = 
(select cpo.id from prdcat.customer_product_Offer cpo
where cpo.customer_ref = 54765460
and cpo.product_offer_id = 253)
And Upper(ORDER_AIM) = Upper('Connect'))
and cpo1.Is_CSE = 'Y'
and cpo1.B2B_Enabled = 'Y')


--Query to Fetch HO
-- Input : TLC and RSP Customer Name
Select pi.external_ref from prdinv.product_instance pi
where Customer_ref = (select cm.customer_ref from cmown.customer_manager cm
where upper(Customer_Name) like upper('%Airnet%'))
and INSTR ((select Point_Of_Interconnect from NAOWN.NADBGE_2@PRODDATANA where TLC = '100553225'),location_ref_A) <> 0
and pi.product_offer_id = '1'
and Rownum <=1


-- Query to fetch Order History
select  oeh.*
from    OM.CUSTOMER_ORDER CO
                inner join
        OM.CUSTOMER_ORDER_STATE_SUBSTATE COSS
            on COSS.id = CO.CUSTOMER_ORDER_STATE_SUBSTT_ID
                inner join
        OM.CUSTOMER_ORDER_ITEM COI
            on coi.customer_order_id = co.id
                inner join
        PRDCAT.PRODUCT_OFFER PO
            on PO.id = COI.PRODUCT_OFFER_ID
             inner  join
        OM.ORCHESTRATABLE_ENTITY_HISTORY OEH
            on OEH.ORCHESTRATABLE_ENTITY_ID = co.id
--where   coss.state <> 'Closed'
Where CO.EXT_CUSTOMER_ORDER_ID_REF = '100037459'
order by oeh.event_date;


-- Fetch Oorder ITem ID from the Order ID



select Received_Timestamp,acknowledged_timestamp,Status,Request_localpart,Priority from b2bown.messages m
where m.customer_Ref = 54530114
order by m.Received_Timestamp Desc




