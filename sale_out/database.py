import csv
import sqlite3
import pandas as pd
from pandas import Series


import jaydebeapi
import pytest
import xlwt


class Tertiary_download_structure:
    def __init__(self,item_kpi,year,month,brand,weight_pen,sro, pen, quantity, amount_euro, weight_sro):
        self.weight_sro = weight_sro
        self.amount_euro = amount_euro
        self.quantity = quantity
        self.pen = pen
        self.sro = sro
        self.weight_pen = weight_pen
        self.brand = brand
        self.month = month
        self.year = year
        self.item_kpi = item_kpi

class Secondary_total_2021:
    def __init__(self,item_quadra,sales_euro):
        self.sales_euro = sales_euro
        self.item_quadra = item_quadra

class Upload_2021_base_from_quadra_for_daily_totals_distr:
    def __init__(self,year,month,week,sales_method,distributor_name, item_quadra,sale_in_quantity,sales_euro_):

        self.sales_euro_ = sales_euro_
        self.sale_in_quantity = sale_in_quantity
        self.item_quadra = item_quadra
        self.distributor_name = distributor_name
        self.sales_method = sales_method
        self.week = week
        self.month = month
        self.year = year

class Quadra_direct_629():
    def __init__(self,year,month,ff_region,country_region,city_town,organization_name,organization_adress,sales_method,
                 product_code,item_quadra,organization_etalon_id,organization_etalon_name,distributor_etalon_name,
                 distributor_name,distributor_okpo,sales_euro_,promotion,organization_type,organization_status,
                 etalon_code_okpo,delivery_date,office_head_organization,head_office_okpo,quarter_year,half_year,
                 annual_sales_category,med_representative_name,kam_name,week,territory_name,brik_name,sale_in_quantity):
        self.sale_in_quantity = sale_in_quantity
        self.brik_name = brik_name
        self.territory_name = territory_name
        self.week = week
        self.kam_name = kam_name
        self.med_representative_name = med_representative_name
        self.annual_sales_category = annual_sales_category
        self.half_year = half_year
        self.quarter_year = quarter_year
        self.head_office_okpo = head_office_okpo
        self.office_head_organization = office_head_organization
        self.delivery_date = delivery_date
        self.etalon_code_okpo = etalon_code_okpo
        self.organization_status = organization_status
        self.organization_type = organization_type
        self.promotion = promotion
        self.sales_euro_ = sales_euro_
        self.distributor_okpo = distributor_okpo
        self.distributor_name = distributor_name
        self.distributor_etalon_name = distributor_etalon_name
        self.organization_etalon_name = organization_etalon_name
        self.organization_etalon_id = organization_etalon_id
        self.item_quadra = item_quadra
        self.product_code = product_code
        self.sales_method = sales_method
        self.organization_adress = organization_adress
        self.organization_name = organization_name
        self.city_town = city_town
        self.country_region = country_region
        self.ff_region = ff_region
        self.month = month
        self.year = year

class CBase_2021_quadra_workout:
    def classify_base_2021(self,base_2021):
        base_2021_classifyed = []
        for i in base_2021:
            sale_in_quantity = i[31]
            brik_name = i[30]
            territory_name = i[29]
            week = i[28]
            kam_name = i[27]
            med_representative_name = i[26]
            annual_sales_category = i[25]
            half_year = i[24]
            quarter_year = i[23]
            head_office_okpo = i[22]
            office_head_organization = i[21]
            delivery_date = i[20]
            etalon_code_okpo = i[19]
            organization_status = i[18]
            organization_type = i[17]
            promotion = i[16]
            sales_euro_ = i[15]
            distributor_okpo = i[14]
            distributor_name = i[13]
            distributor_etalon_name = i[12]
            organization_etalon_name = i[11]
            organization_etalon_id = i[10]
            item_quadra = i[9]
            product_code = i[8]
            sales_method = i[7]
            organization_adress = i[6]
            organization_name = i[5]
            city_town = i[4]
            country_region = i[3]
            ff_region = i[2]
            month = i[1]
            year = i[0]
            entry = Quadra_direct_629(year,month,ff_region,country_region,city_town,organization_name,organization_adress,sales_method,
                     product_code,item_quadra,organization_etalon_id,organization_etalon_name,distributor_etalon_name,
                     distributor_name,distributor_okpo,sales_euro_,promotion,organization_type,organization_status,
                     etalon_code_okpo,delivery_date,office_head_organization,head_office_okpo,quarter_year,half_year,
                     annual_sales_category,med_representative_name,kam_name,week,territory_name,brik_name,sale_in_quantity)
            base_2021_classifyed.append(entry)
        return base_2021_classifyed

    def upload_2021_base_from_quadra(self):
        try:
            # jTDS Driver.
            driver_name = "net.sourceforge.jtds.jdbc.Driver"

            # jTDS Connection string.
            connection_url = "jdbc:jtds:sqlserver://62.149.15.123:1433/medowl_grindex"

            # jTDS Connection properties.
            # Some additional connection properties you may want to use
            # "domain": "<domain>"
            # "ssl": "require"
            # "useNTLMv2": "true"
            # See the FAQ for details http://jtds.sourceforge.net/faq.html
            connection_properties = {
                "user": "grindex",
                "password": "xednirg",
            }

            # Path to jTDS Jar
            jar_path = 'C:\jtds-1.3.1-dist\\jtds-1.3.1.jar'

            # Establish connection.
            connection = jaydebeapi.connect(driver_name, connection_url, connection_properties, jar_path)
            cursor = connection.cursor()

            # Execute test query.
            cursor.execute("SELECT orgSalesByProduct.year AS year,"
                           + " month.name AS month,"
                           + " region.name AS ff_region,"
                           + " province.name AS country_region,"
                           + " town.name AS city_town,"
                           + " organization.name AS organization_name,"
                           + " organization.postal_address AS organization_adress,"
                           + " category.name AS sales_method,"
                           + " product.code AS product_code,"
                           + " product.name AS item_quadra,"
                           + " orgSalesByProduct.organization_etalon_id AS organization_etalon_id,"
                           + " orgSalesByProduct.organization_etalon_name AS organization_etalon_name,"
                           + " orgSalesByProduct.distributor_etalon_name AS distributor_etalon_name,"
                           + " distributor.name AS distributor_name,"
                           + " orgSalesByProduct.distributor_okpo AS distributor_okpo,"
                           + " productPriceDynamics.price_fact * ROUND(orgSalesByProduct.quantity, 2) AS sales_euro_,"
                           + " productCategory2.name AS promotion,"
                           + " organizationType.name AS organization_type,"
                           + " organizationStatus.name AS organization_status,"
                           + " orgSalesByProduct.etalon_code_okpo AS etalon_code_okpo,"
                           + " orgSalesByProduct.delivery_date AS delivery_date,"
                           + " 	   "
                           + " ("
                           + " SELECT MAX(orgOf.name+' '+sType.name+' '+orgOf.street_name+', '+orgOf.street_number)"
                           + " FROM organization AS orgOf"
                           + "     LEFT JOIN street_type sType ON sType.id = orgOf.street_type_id"
                           + " WHERE orgOf.etalon_id = orgSalesByProduct.head_organization_etalon_id"
                           + " ) AS office_head_organization,"
                           + " CASE"
                           + "     WHEN orgSalesByProduct.head_office_okpo IS NOT NULL"
                           + "     THEN orgSalesByProduct.head_office_okpo"
                           + "     ELSE ''"
                           + " END AS head_office_okpo,"
                           + " CASE"
                           + "     WHEN month.id = 1"
                           + "         OR month.id = 2"
                           + "         OR month.id = 3"
                           + "     THEN '1-й квартал'"
                           + "     ELSE CASE"
                           + "             WHEN month.id = 4"
                           + "                     OR month.id = 5"
                           + "                     OR month.id = 6"
                           + "             THEN '2-й квартал'"
                           + "             ELSE CASE"
                           + "                         WHEN month.id = 7"
                           + "                             OR month.id = 8"
                           + "                             OR month.id = 9"
                           + "                         THEN '3-й квартал'"
                           + "                         ELSE CASE"
                           + "                                 WHEN month.id = 10"
                           + "                                     OR month.id = 11"
                           + "                                     OR month.id = 12"
                           + "                                 THEN '4-й квартал'"
                           + "                                 ELSE ''"
                           + "                             END"
                           + "                     END"
                           + "         END"
                           + " END AS quarter_year,"
                           + " "
                           + " CASE"
                           + "     WHEN month.id IN(1, 2, 3, 4, 5, 6)"
                           + "     THEN '1-е полугодие'"
                           + "     ELSE '2-е полугодие'"
                           + " END AS half_year,"
                           + " "
                           + " ("
                           + " SELECT sbm.name"
                           + " FROM sales_by_month_category AS sbm"
                           + " WHERE sbm.pos_status = 0"
                           + "     AND sbm.id ="
                           + " ("
                           + " SELECT orgfet.sales_by_month_category_id"
                           + " FROM organization_from_etalon AS orgfet"
                           + " WHERE orgfet.pos_status = 0"
                           + "         AND orgfet.etalon_id = organization.etalon_id"
                           + " )"
                           + " ) AS annual_sales_category,"
                           + " "
                           + " "
                           + " "
                           + " CASE"
                           + "     WHEN empl.id IS NOT NULL"
                           + "         AND emplSG.role_id = 3"
                           + "         AND empl.status_id = 1"
                           + "     THEN empl.last_name+' '+empl.first_name"
                           + "     ELSE ''"
                           + " END AS med_representative_name,"
                           + " CASE"
                           + "     WHEN kam.id IS NOT NULL"
                           + "     THEN kam.last_name+' '+kam.first_name"
                           + "     ELSE ''"
                           + " END AS kam_name,"
                           + " DATEPART(ISO_WEEK, orgSalesByProduct.delivery_date) AS week,"
                           + " territory.name AS territory_name,"
                           + " brick.name AS brik_name,"
                           + " ROUND(orgSalesByProduct.quantity, 2) AS sale_in_quantity"
                           + " "
                           + " "
                           + " "
                           + " FROM organization_sales_by_product AS orgSalesByProduct"
                           + " LEFT JOIN month_classifier month ON orgSalesByProduct.month_id = month.id"
                           + " LEFT JOIN organization organization ON organization.id = orgSalesByProduct.organization_id"
                           + " LEFT JOIN registration_type registration ON organization.registration_type_id = registration.id"
                           + " LEFT JOIN rayon rayon ON organization.rayon_id = rayon.id"
                           + " LEFT JOIN town town ON town.id = orgSalesByProduct.town_id"
                           + " LEFT JOIN oblast province ON province.id = town.oblast_id"
                           + " LEFT JOIN region region ON region.id = province.region_id"
                           + " LEFT JOIN region region2 ON region2.id = province.region2_id"
                           + " LEFT JOIN region region3 ON region3.id = province.region3_id"
                           + " LEFT JOIN region region4 ON region4.id = province.region4_id"
                           + " LEFT JOIN oblast_rayon admRayon ON admRayon.id = town.oblast_rayon_id"
                           + " LEFT JOIN organization_type organizationType ON organizationType.id = organization.organization_type_id"
                           + " LEFT JOIN organization_status organizationStatus ON organizationStatus.id = organizationType.organization_status_id"
                           + " LEFT JOIN product product ON product.id = orgSalesByProduct.product_id"
                           + " LEFT JOIN product_series series ON series.id = product.product_series_id"
                           + " LEFT JOIN product_series_group psGroup ON psGroup.id = series.product_series_group_id"
                           + " LEFT JOIN product_category category ON category.id = product.product_category_id"
                           + " LEFT JOIN product_category2 category2 ON category2.id = product.product_category2_id"
                           + " LEFT JOIN product_category_atc productCategoryAtc ON productCategoryAtc.id = product.product_category_atc_id"
                           + " LEFT JOIN employee empl ON empl.id = orgSalesByProduct.employee_id"
                           + " LEFT JOIN security_group emplSG ON emplSG.id = empl.security_group_id"
                           + " LEFT JOIN organization_drugstore_office odo ON odo.organization_id = orgSalesByProduct.head_organization_id"
                           + "                                             AND odo.pos_status = 0"
                           + "                                             AND odo.active = 1"
                           + " LEFT JOIN employee kam ON kam.id = odo.manager_id"
                           + " LEFT JOIN organization_sales_by_product_op osbop ON osbop.organization_etalon_id = organization.etalon_id"
                           + "                                                     AND osbop.product_id = product.id"
                           + " LEFT JOIN territory_brick brick ON osbop.territory_brick_id = brick.id"
                           + " LEFT JOIN territory territory ON brick.territory_id = territory.id"
                           + " LEFT JOIN town headTown ON orgSalesByProduct.head_organization_town_id = headTown.id"
                           + " LEFT JOIN oblast headObl ON headTown.oblast_id = headObl.id"
                           + " LEFT JOIN organization distributor ON orgSalesByProduct.distributor_id = distributor.id"
                           + " LEFT JOIN sales_by_month_category salesByMonthCategory ON salesByMonthCategory.id = organization.sales_by_month_category_id"
                           + " LEFT JOIN category orgCat ON orgCat.id = organization.organization_category_id"
                           + " LEFT JOIN product_price_dynamics productPriceDynamics ON productPriceDynamics.pos_status = 0"
                           + "                                                         AND productPriceDynamics.month_id = orgSalesByProduct.month_id"
                           + "                                                         AND productPriceDynamics.year = orgSalesByProduct.year"
                           + "                                                         AND productPriceDynamics.product_id = orgSalesByProduct.product_id"
                           + " LEFT JOIN product_category2 productCategory2 ON productPriceDynamics.product_category2_id = productCategory2.id"
                           + " LEFT JOIN sales_type salesType ON organization.id = salesType.organization_id"
                           + "                                 AND salesType.year = orgSalesByProduct.year"
                           + "                                 AND salesType.month_id = month.id"
                           + "                                 AND salesType.pos_status = 0"
                           + " LEFT JOIN organization_specialization orgSpecialization ON organization.organization_specialization_id = orgSpecialization.id"
                           + " WHERE orgSalesByProduct.year = 2021")
            res = cursor.fetchall()

            x = CBase_2021_quadra_workout()
            base_2021_classifyed = x.classify_base_2021(res)
            return base_2021_classifyed
        except Exception as err:
            print(str(err))
    def upload_2021_base_from_quadra_unclass(self):
        try:
            # jTDS Driver.
            driver_name = "net.sourceforge.jtds.jdbc.Driver"

            # jTDS Connection string.
            connection_url = "jdbc:jtds:sqlserver://62.149.15.123:1433/medowl_grindex"

            # jTDS Connection properties.
            # Some additional connection properties you may want to use
            # "domain": "<domain>"
            # "ssl": "require"
            # "useNTLMv2": "true"
            # See the FAQ for details http://jtds.sourceforge.net/faq.html
            connection_properties = {
                "user": "grindex",
                "password": "xednirg",
            }

            # Path to jTDS Jar
            jar_path = 'C:\jtds-1.3.1-dist\\jtds-1.3.1.jar'

            # Establish connection.
            connection = jaydebeapi.connect(driver_name, connection_url, connection_properties, jar_path)
            cursor = connection.cursor()

            # Execute test query.
            cursor.execute("SELECT orgSalesByProduct.year AS year,"
                           + " month.name AS month,"
                           + " region.name AS ff_region,"
                           + " province.name AS country_region,"
                           + " town.name AS city_town,"
                           + " organization.name AS organization_name,"
                           + " organization.postal_address AS organization_adress,"
                           + " category.name AS sales_method,"
                           + " product.code AS product_code,"
                           + " product.name AS item_quadra,"
                           + " orgSalesByProduct.organization_etalon_id AS organization_etalon_id,"
                           + " orgSalesByProduct.organization_etalon_name AS organization_etalon_name,"
                           + " orgSalesByProduct.distributor_etalon_name AS distributor_etalon_name,"
                           + " distributor.name AS distributor_name,"
                           + " orgSalesByProduct.distributor_okpo AS distributor_okpo,"
                           + " productPriceDynamics.price_fact * ROUND(orgSalesByProduct.quantity, 2) AS sales_euro_,"
                           + " productCategory2.name AS promotion,"
                           + " organizationType.name AS organization_type,"
                           + " organizationStatus.name AS organization_status,"
                           + " orgSalesByProduct.etalon_code_okpo AS etalon_code_okpo,"
                           + " orgSalesByProduct.delivery_date AS delivery_date,"
                           + " 	   "
                           + " ("
                           + " SELECT MAX(orgOf.name+' '+sType.name+' '+orgOf.street_name+', '+orgOf.street_number)"
                           + " FROM organization AS orgOf"
                           + "     LEFT JOIN street_type sType ON sType.id = orgOf.street_type_id"
                           + " WHERE orgOf.etalon_id = orgSalesByProduct.head_organization_etalon_id"
                           + " ) AS office_head_organization,"
                           + " CASE"
                           + "     WHEN orgSalesByProduct.head_office_okpo IS NOT NULL"
                           + "     THEN orgSalesByProduct.head_office_okpo"
                           + "     ELSE ''"
                           + " END AS head_office_okpo,"
                           + " CASE"
                           + "     WHEN month.id = 1"
                           + "         OR month.id = 2"
                           + "         OR month.id = 3"
                           + "     THEN '1-й квартал'"
                           + "     ELSE CASE"
                           + "             WHEN month.id = 4"
                           + "                     OR month.id = 5"
                           + "                     OR month.id = 6"
                           + "             THEN '2-й квартал'"
                           + "             ELSE CASE"
                           + "                         WHEN month.id = 7"
                           + "                             OR month.id = 8"
                           + "                             OR month.id = 9"
                           + "                         THEN '3-й квартал'"
                           + "                         ELSE CASE"
                           + "                                 WHEN month.id = 10"
                           + "                                     OR month.id = 11"
                           + "                                     OR month.id = 12"
                           + "                                 THEN '4-й квартал'"
                           + "                                 ELSE ''"
                           + "                             END"
                           + "                     END"
                           + "         END"
                           + " END AS quarter_year,"
                           + " "
                           + " CASE"
                           + "     WHEN month.id IN(1, 2, 3, 4, 5, 6)"
                           + "     THEN '1-е полугодие'"
                           + "     ELSE '2-е полугодие'"
                           + " END AS half_year,"
                           + " "
                           + " ("
                           + " SELECT sbm.name"
                           + " FROM sales_by_month_category AS sbm"
                           + " WHERE sbm.pos_status = 0"
                           + "     AND sbm.id ="
                           + " ("
                           + " SELECT orgfet.sales_by_month_category_id"
                           + " FROM organization_from_etalon AS orgfet"
                           + " WHERE orgfet.pos_status = 0"
                           + "         AND orgfet.etalon_id = organization.etalon_id"
                           + " )"
                           + " ) AS annual_sales_category,"
                           + " "
                           + " "
                           + " "
                           + " CASE"
                           + "     WHEN empl.id IS NOT NULL"
                           + "         AND emplSG.role_id = 3"
                           + "         AND empl.status_id = 1"
                           + "     THEN empl.last_name+' '+empl.first_name"
                           + "     ELSE ''"
                           + " END AS med_representative_name,"
                           + " CASE"
                           + "     WHEN kam.id IS NOT NULL"
                           + "     THEN kam.last_name+' '+kam.first_name"
                           + "     ELSE ''"
                           + " END AS kam_name,"
                           + " DATEPART(ISO_WEEK, orgSalesByProduct.delivery_date) AS week,"
                           + " territory.name AS territory_name,"
                           + " brick.name AS brik_name,"
                           + " ROUND(orgSalesByProduct.quantity, 2) AS sale_in_quantity"
                           + " "
                           + " "
                           + " "
                           + " FROM organization_sales_by_product AS orgSalesByProduct"
                           + " LEFT JOIN month_classifier month ON orgSalesByProduct.month_id = month.id"
                           + " LEFT JOIN organization organization ON organization.id = orgSalesByProduct.organization_id"
                           + " LEFT JOIN registration_type registration ON organization.registration_type_id = registration.id"
                           + " LEFT JOIN rayon rayon ON organization.rayon_id = rayon.id"
                           + " LEFT JOIN town town ON town.id = orgSalesByProduct.town_id"
                           + " LEFT JOIN oblast province ON province.id = town.oblast_id"
                           + " LEFT JOIN region region ON region.id = province.region_id"
                           + " LEFT JOIN region region2 ON region2.id = province.region2_id"
                           + " LEFT JOIN region region3 ON region3.id = province.region3_id"
                           + " LEFT JOIN region region4 ON region4.id = province.region4_id"
                           + " LEFT JOIN oblast_rayon admRayon ON admRayon.id = town.oblast_rayon_id"
                           + " LEFT JOIN organization_type organizationType ON organizationType.id = organization.organization_type_id"
                           + " LEFT JOIN organization_status organizationStatus ON organizationStatus.id = organizationType.organization_status_id"
                           + " LEFT JOIN product product ON product.id = orgSalesByProduct.product_id"
                           + " LEFT JOIN product_series series ON series.id = product.product_series_id"
                           + " LEFT JOIN product_series_group psGroup ON psGroup.id = series.product_series_group_id"
                           + " LEFT JOIN product_category category ON category.id = product.product_category_id"
                           + " LEFT JOIN product_category2 category2 ON category2.id = product.product_category2_id"
                           + " LEFT JOIN product_category_atc productCategoryAtc ON productCategoryAtc.id = product.product_category_atc_id"
                           + " LEFT JOIN employee empl ON empl.id = orgSalesByProduct.employee_id"
                           + " LEFT JOIN security_group emplSG ON emplSG.id = empl.security_group_id"
                           + " LEFT JOIN organization_drugstore_office odo ON odo.organization_id = orgSalesByProduct.head_organization_id"
                           + "                                             AND odo.pos_status = 0"
                           + "                                             AND odo.active = 1"
                           + " LEFT JOIN employee kam ON kam.id = odo.manager_id"
                           + " LEFT JOIN organization_sales_by_product_op osbop ON osbop.organization_etalon_id = organization.etalon_id"
                           + "                                                     AND osbop.product_id = product.id"
                           + " LEFT JOIN territory_brick brick ON osbop.territory_brick_id = brick.id"
                           + " LEFT JOIN territory territory ON brick.territory_id = territory.id"
                           + " LEFT JOIN town headTown ON orgSalesByProduct.head_organization_town_id = headTown.id"
                           + " LEFT JOIN oblast headObl ON headTown.oblast_id = headObl.id"
                           + " LEFT JOIN organization distributor ON orgSalesByProduct.distributor_id = distributor.id"
                           + " LEFT JOIN sales_by_month_category salesByMonthCategory ON salesByMonthCategory.id = organization.sales_by_month_category_id"
                           + " LEFT JOIN category orgCat ON orgCat.id = organization.organization_category_id"
                           + " LEFT JOIN product_price_dynamics productPriceDynamics ON productPriceDynamics.pos_status = 0"
                           + "                                                         AND productPriceDynamics.month_id = orgSalesByProduct.month_id"
                           + "                                                         AND productPriceDynamics.year = orgSalesByProduct.year"
                           + "                                                         AND productPriceDynamics.product_id = orgSalesByProduct.product_id"
                           + " LEFT JOIN product_category2 productCategory2 ON productPriceDynamics.product_category2_id = productCategory2.id"
                           + " LEFT JOIN sales_type salesType ON organization.id = salesType.organization_id"
                           + "                                 AND salesType.year = orgSalesByProduct.year"
                           + "                                 AND salesType.month_id = month.id"
                           + "                                 AND salesType.pos_status = 0"
                           + " LEFT JOIN organization_specialization orgSpecialization ON organization.organization_specialization_id = orgSpecialization.id"
                           + " WHERE orgSalesByProduct.year = 2021")
            res = cursor.fetchall()

            return res
        except Exception as err:
            print(str(err))

    def save_base_2021_to_csv(self):
        global base_2021_classifyed_, item
        filename = 'base_2021.csv'
        final_list = []
        try:
            # jTDS Driver.
            driver_name = "net.sourceforge.jtds.jdbc.Driver"

            # jTDS Connection string.
            connection_url = "jdbc:jtds:sqlserver://62.149.15.123:1433/medowl_grindex"

            # jTDS Connection properties.
            # Some additional connection properties you may want to use
            # "domain": "<domain>"
            # "ssl": "require"
            # "useNTLMv2": "true"
            # See the FAQ for details http://jtds.sourceforge.net/faq.html
            connection_properties = {
                "user": "grindex",
                "password": "xednirg",
            }

            # Path to jTDS Jar
            jar_path = 'C:\jtds-1.3.1-dist\\jtds-1.3.1.jar'

            # Establish connection.
            connection = jaydebeapi.connect(driver_name, connection_url, connection_properties, jar_path)
            cursor = connection.cursor()

            # Execute test query.
            cursor.execute("SELECT orgSalesByProduct.year AS year,"
                           + " month.name AS month,"
                           + " region.name AS ff_region,"
                           + " province.name AS country_region,"
                           + " town.name AS city_town,"
                           + " organization.name AS organization_name,"
                           + " organization.postal_address AS organization_adress,"
                           + " category.name AS sales_method,"
                           + " product.code AS product_code,"
                           + " product.name AS item_quadra,"
                           + " orgSalesByProduct.organization_etalon_id AS organization_etalon_id,"
                           + " orgSalesByProduct.organization_etalon_name AS organization_etalon_name,"
                           + " orgSalesByProduct.distributor_etalon_name AS distributor_etalon_name,"
                           + " distributor.name AS distributor_name,"
                           + " orgSalesByProduct.distributor_okpo AS distributor_okpo,"
                           + " productPriceDynamics.price_fact * ROUND(orgSalesByProduct.quantity, 2) AS sales_euro_,"
                           + " productCategory2.name AS promotion,"
                           + " organizationType.name AS organization_type,"
                           + " organizationStatus.name AS organization_status,"
                           + " orgSalesByProduct.etalon_code_okpo AS etalon_code_okpo,"
                           + " orgSalesByProduct.delivery_date AS delivery_date,"
                           + " 	   "
                           + " ("
                           + " SELECT MAX(orgOf.name+' '+sType.name+' '+orgOf.street_name+', '+orgOf.street_number)"
                           + " FROM organization AS orgOf"
                           + "     LEFT JOIN street_type sType ON sType.id = orgOf.street_type_id"
                           + " WHERE orgOf.etalon_id = orgSalesByProduct.head_organization_etalon_id"
                           + " ) AS office_head_organization,"
                           + " CASE"
                           + "     WHEN orgSalesByProduct.head_office_okpo IS NOT NULL"
                           + "     THEN orgSalesByProduct.head_office_okpo"
                           + "     ELSE ''"
                           + " END AS head_office_okpo,"
                           + " CASE"
                           + "     WHEN month.id = 1"
                           + "         OR month.id = 2"
                           + "         OR month.id = 3"
                           + "     THEN '1-й квартал'"
                           + "     ELSE CASE"
                           + "             WHEN month.id = 4"
                           + "                     OR month.id = 5"
                           + "                     OR month.id = 6"
                           + "             THEN '2-й квартал'"
                           + "             ELSE CASE"
                           + "                         WHEN month.id = 7"
                           + "                             OR month.id = 8"
                           + "                             OR month.id = 9"
                           + "                         THEN '3-й квартал'"
                           + "                         ELSE CASE"
                           + "                                 WHEN month.id = 10"
                           + "                                     OR month.id = 11"
                           + "                                     OR month.id = 12"
                           + "                                 THEN '4-й квартал'"
                           + "                                 ELSE ''"
                           + "                             END"
                           + "                     END"
                           + "         END"
                           + " END AS quarter_year,"
                           + " "
                           + " CASE"
                           + "     WHEN month.id IN(1, 2, 3, 4, 5, 6)"
                           + "     THEN '1-е полугодие'"
                           + "     ELSE '2-е полугодие'"
                           + " END AS half_year,"
                           + " "
                           + " ("
                           + " SELECT sbm.name"
                           + " FROM sales_by_month_category AS sbm"
                           + " WHERE sbm.pos_status = 0"
                           + "     AND sbm.id ="
                           + " ("
                           + " SELECT orgfet.sales_by_month_category_id"
                           + " FROM organization_from_etalon AS orgfet"
                           + " WHERE orgfet.pos_status = 0"
                           + "         AND orgfet.etalon_id = organization.etalon_id"
                           + " )"
                           + " ) AS annual_sales_category,"
                           + " "
                           + " "
                           + " "
                           + " CASE"
                           + "     WHEN empl.id IS NOT NULL"
                           + "         AND emplSG.role_id = 3"
                           + "         AND empl.status_id = 1"
                           + "     THEN empl.last_name+' '+empl.first_name"
                           + "     ELSE ''"
                           + " END AS med_representative_name,"
                           + " CASE"
                           + "     WHEN kam.id IS NOT NULL"
                           + "     THEN kam.last_name+' '+kam.first_name"
                           + "     ELSE ''"
                           + " END AS kam_name,"
                           + " DATEPART(ISO_WEEK, orgSalesByProduct.delivery_date) AS week,"
                           + " territory.name AS territory_name,"
                           + " brick.name AS brik_name,"
                           + " ROUND(orgSalesByProduct.quantity, 2) AS sale_in_quantity"
                           + " "
                           + " "
                           + " "
                           + " FROM organization_sales_by_product AS orgSalesByProduct"
                           + " LEFT JOIN month_classifier month ON orgSalesByProduct.month_id = month.id"
                           + " LEFT JOIN organization organization ON organization.id = orgSalesByProduct.organization_id"
                           + " LEFT JOIN registration_type registration ON organization.registration_type_id = registration.id"
                           + " LEFT JOIN rayon rayon ON organization.rayon_id = rayon.id"
                           + " LEFT JOIN town town ON town.id = orgSalesByProduct.town_id"
                           + " LEFT JOIN oblast province ON province.id = town.oblast_id"
                           + " LEFT JOIN region region ON region.id = province.region_id"
                           + " LEFT JOIN region region2 ON region2.id = province.region2_id"
                           + " LEFT JOIN region region3 ON region3.id = province.region3_id"
                           + " LEFT JOIN region region4 ON region4.id = province.region4_id"
                           + " LEFT JOIN oblast_rayon admRayon ON admRayon.id = town.oblast_rayon_id"
                           + " LEFT JOIN organization_type organizationType ON organizationType.id = organization.organization_type_id"
                           + " LEFT JOIN organization_status organizationStatus ON organizationStatus.id = organizationType.organization_status_id"
                           + " LEFT JOIN product product ON product.id = orgSalesByProduct.product_id"
                           + " LEFT JOIN product_series series ON series.id = product.product_series_id"
                           + " LEFT JOIN product_series_group psGroup ON psGroup.id = series.product_series_group_id"
                           + " LEFT JOIN product_category category ON category.id = product.product_category_id"
                           + " LEFT JOIN product_category2 category2 ON category2.id = product.product_category2_id"
                           + " LEFT JOIN product_category_atc productCategoryAtc ON productCategoryAtc.id = product.product_category_atc_id"
                           + " LEFT JOIN employee empl ON empl.id = orgSalesByProduct.employee_id"
                           + " LEFT JOIN security_group emplSG ON emplSG.id = empl.security_group_id"
                           + " LEFT JOIN organization_drugstore_office odo ON odo.organization_id = orgSalesByProduct.head_organization_id"
                           + "                                             AND odo.pos_status = 0"
                           + "                                             AND odo.active = 1"
                           + " LEFT JOIN employee kam ON kam.id = odo.manager_id"
                           + " LEFT JOIN organization_sales_by_product_op osbop ON osbop.organization_etalon_id = organization.etalon_id"
                           + "                                                     AND osbop.product_id = product.id"
                           + " LEFT JOIN territory_brick brick ON osbop.territory_brick_id = brick.id"
                           + " LEFT JOIN territory territory ON brick.territory_id = territory.id"
                           + " LEFT JOIN town headTown ON orgSalesByProduct.head_organization_town_id = headTown.id"
                           + " LEFT JOIN oblast headObl ON headTown.oblast_id = headObl.id"
                           + " LEFT JOIN organization distributor ON orgSalesByProduct.distributor_id = distributor.id"
                           + " LEFT JOIN sales_by_month_category salesByMonthCategory ON salesByMonthCategory.id = organization.sales_by_month_category_id"
                           + " LEFT JOIN category orgCat ON orgCat.id = organization.organization_category_id"
                           + " LEFT JOIN product_price_dynamics productPriceDynamics ON productPriceDynamics.pos_status = 0"
                           + "                                                         AND productPriceDynamics.month_id = orgSalesByProduct.month_id"
                           + "                                                         AND productPriceDynamics.year = orgSalesByProduct.year"
                           + "                                                         AND productPriceDynamics.product_id = orgSalesByProduct.product_id"
                           + " LEFT JOIN product_category2 productCategory2 ON productPriceDynamics.product_category2_id = productCategory2.id"
                           + " LEFT JOIN sales_type salesType ON organization.id = salesType.organization_id"
                           + "                                 AND salesType.year = orgSalesByProduct.year"
                           + "                                 AND salesType.month_id = month.id"
                           + "                                 AND salesType.pos_status = 0"
                           + " LEFT JOIN organization_specialization orgSpecialization ON organization.organization_specialization_id = orgSpecialization.id"
                           + " WHERE orgSalesByProduct.year = 2021")
            res = cursor.fetchall()

            x = CBase_2021_quadra_workout()
            base_2021_classifyed_ = x.classify_base_2021(res)
        except Exception as err:
            print(str(err))
        sum_tot_euro = 0
        with open('work_base.csv', "w", newline="", encoding='utf-16') as file:
            columns = ['year', 'month', 'ff_region', 'country_region', 'city_town', 'organization_name',
                       'organization_adress', 'sales_method',
                       'product_code', 'item_quadra', 'organization_etalon_id', 'organization_etalon_name',
                       'distributor_etalon_name',
                       'distributor_name', 'distributor_okpo', 'sales_euro_', 'promotion', 'organization_type',
                       'organization_status',
                       'etalon_code_okpo', 'delivery_date', 'office_head_organization', 'head_office_okpo',
                       'quarter_year', 'half_year',
                       'annual_sales_category', 'med_representative_name', 'kam_name', 'week', 'territory_name',
                       'brik_name', 'sale_in_quantity']
            writer = csv.writer(file)
            list_base_2021 = [['year', 'month', 'ff_region', 'country_region', 'city_town', 'organization_name',
                       'organization_adress', 'sales_method',
                       'product_code', 'item_quadra', 'organization_etalon_id', 'organization_etalon_name',
                       'distributor_etalon_name',
                       'distributor_name', 'distributor_okpo', 'sales_euro_', 'promotion', 'organization_type',
                       'organization_status',
                       'etalon_code_okpo', 'delivery_date', 'office_head_organization', 'head_office_okpo',
                       'quarter_year', 'half_year',
                       'annual_sales_category', 'med_representative_name', 'kam_name', 'week', 'territory_name',
                       'brik_name', 'sale_in_quantity']]
            for item in base_2021_classifyed_:
                z0 = str(item.year).replace(",", "")
                y1 = str(item.sales_euro_).replace(".", ",")
                y2 = str(item.sale_in_quantity).replace(".", ",")
                item_ = [[str(z0), str(item.month), str(item.ff_region),
                         str(item.country_region), str(item.city_town),
                         str(item.organization_name), str(item.organization_adress),
                         str(item.sales_method),
                         str(item.product_code), str(item.item_quadra),
                         str(item.organization_etalon_id),
                         str(item.organization_etalon_name),
                         str(item.distributor_etalon_name),
                         str(item.distributor_name), str(item.distributor_okpo),
                         str(y1), str(item.promotion), str(item.organization_type),
                         str(item.organization_status),
                         str(item.etalon_code_okpo), str(item.delivery_date),
                         str(item.office_head_organization),
                         str(item.head_office_okpo), str(item.quarter_year),
                         str(item.half_year),
                         str(item.annual_sales_category),
                         str(item.med_representative_name), str(item.kam_name),
                         str(item.week), str(item.territory_name), str(item.brik_name),
                         str(y2)]]
                list_base_2021.append(item_)



            sum_tot_euro += item.sales_euro_
            writer.writerows(list_base_2021)

        print(sum_tot_euro)

class CTest_SAles_report_classification:
    def __init__(self,year,month,week,distr,item_quadra,sec_quantity,sec_euro):
        self.sec_euro = sec_euro
        self.sec_quantity = sec_quantity
        self.item_quadra = item_quadra
        self.distr = distr
        self.week = week
        self.month = month
        self.year = year



class CTest_SAles_report_creation:
    def test_rep(self,base_raw):
        years = []
        months = []
        items = []
        distributor_name = []
        weeks = []
        secondary_sales_packs = 0
        secondary_sales_euro = 0
        for i in base_raw:
            if i.year not in years:
                years.append(i.year)
            if i.month not in months:
                months.append(i.month)
            if i.item_quadra not in items:
                items.append(i.item_quadra)
            if i.distributor_name not in distributor_name:
                distributor_name.append(i.distributor_name)
            if i.week not in weeks:
                weeks.append(i.week)
        sales_by_product = []
        for i in base_raw:
            for year in years:
                for month in months:
                    for week in weeks:
                        for distr in distributor_name:
                            for item in items:
                                if i.year == year and i.month == month and i.week == week and i.distributor_name == distr and i.item_quadra == item:
                                    sales_by_product.append(CTest_SAles_report_classification(i.year, i.month,i.week,i.distributor_name,i.item_quadra, i.sale_in_quantity, i.sales_euro_))
        return sales_by_product



class CEXtract_database_tertiary:
    def read_item(conn, year):
        tertiary_list = []
        with sqlite3.connect("tertiary_sales_database.db") as conn:
            cursor = conn.cursor()
            cursor.execute(f"SELECT DISTINCT items.item_kpi_report, tertiary_sales.Year, tertiary_sales.PeriodName, items.brand, tertiary_sales.WeightPenetration,  tertiary_sales.SRO, tertiary_sales.Penetration, tertiary_sales.Quantity, tertiary_sales.Volume, tertiary_sales.WeightSRO from items join tertiary_sales on tertiary_sales.Fullmedicationname = items.item_proxima WHERE tertiary_sales.Year = {year}")
            results = cursor.fetchall()
            for i in results:
                y_1 = str(i[4]).replace(',', '.')
                y_2 = str(i[5]).replace(',', '.')
                y_3 = str(i[6]).replace(',', '.')
                y_4 = str(i[7]).replace(',', '.')
                y_5 = str(i[8]).replace(',', '.')
                y_6 = str(i[9]).replace(',', '.')
                z = Tertiary_download_structure(i[0], i[1], i[2], i[3], y_1, y_2, y_3, y_4, y_5,y_6)
                tertiary_list.append(z)
        return tertiary_list


    def read_item_2020_w_commas(conn,year):
        with sqlite3.connect("tertiary_sales_database.db") as conn:
            cursor = conn.cursor()
            cursor.execute(
                f"SELECT DISTINCT items.item_kpi_report, tertiary_sales.Year, tertiary_sales.PeriodName, items.brand, tertiary_sales.WeightPenetration,  tertiary_sales.SRO, tertiary_sales.Penetration, tertiary_sales.Quantity, tertiary_sales.Volume, tertiary_sales.WeightSRO from items join tertiary_sales on tertiary_sales.Fullmedicationname = items.item_proxima WHERE tertiary_sales.Year = {year}")
            results = cursor.fetchall()
            tertiary_list = []
            for i in results:
                z = Tertiary_download_structure(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8],i[9])
                tertiary_list.append(z)
        return tertiary_list

    def save_2020_items_to_csv(self, filename, list_months,year):
        with sqlite3.connect("tertiary_sales_database.db") as conn:
            cursor = conn.cursor()
            cursor.execute(
                f"SELECT DISTINCT items.item_kpi_report, tertiary_sales.Year, tertiary_sales.PeriodName, items.brand, tertiary_sales.WeightPenetration,  tertiary_sales.SRO, tertiary_sales.Penetration, tertiary_sales.Quantity, tertiary_sales.Volume, tertiary_sales.WeightSRO from items join tertiary_sales on tertiary_sales.Fullmedicationname = items.item_proxima WHERE tertiary_sales.Year = {year}")
            results = cursor.fetchall()
            tertiary_list = []
            for i in results:
                y_1 = str(i[4]).replace(',', '.')
                y_2 = str(i[5]).replace(',', '.')
                y_3 = str(i[6]).replace(',', '.')
                y_4 = str(i[7]).replace(',', '.')
                y_5 = str(i[8]).replace(',', '.')
                y_6 = str(i[9]).replace(',', '.')
                z = Tertiary_download_structure(i[0], i[1], i[2], i[3], y_1, y_2, y_3, y_4, y_5,y_6)
                tertiary_list.append(z)
        final_list = []
        for month in list_months:
            for i in tertiary_list:
                if month == i.month:
                    final_list.append([{"year": str(i.year), "month": i.month, "brand": i.brand, "item": i.item_kpi,
                                        "weight_penetration": str(i.weight_pen),
                                        "sro": str(i.sro), "penetration": str(i.pen),
                                        "quantity": str(i.quantity), "amount_euro": str(i.amount_euro),"weighted_sro":str(i.weight_sro)}])
        with open(filename, "w", newline="", encoding='utf-16') as file:
            columns = ["year","month","brand","item","weight_penetration","sro","penetration","quantity","amount_euro","weighted_sro"]
            writer = csv.DictWriter(file, fieldnames=columns)
            writer.writeheader()
            for item in final_list:
                writer.writerows(item)

    def save_items_2020_to_csv_with_commas(self, filename, list_months,year):
        with sqlite3.connect("tertiary_sales_database.db") as conn:
            cursor = conn.cursor()
            cursor.execute(
                f"SELECT DISTINCT items.item_kpi_report, tertiary_sales.Year, tertiary_sales.PeriodName, items.brand, tertiary_sales.WeightPenetration,  tertiary_sales.SRO, tertiary_sales.Penetration, tertiary_sales.Quantity, tertiary_sales.Volume, tertiary_sales.WeightSRO from items join tertiary_sales on tertiary_sales.Fullmedicationname = items.item_proxima WHERE tertiary_sales.Year = {year}")
            results = cursor.fetchall()
            tertiary_list = []
            for i in results:

                z = Tertiary_download_structure(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8],i[9])
                tertiary_list.append(z)
        final_list = []
        for month in list_months:
            for i in tertiary_list:
                if month == i.month:
                    final_list.append([{"year": str(i.year), "month": i.month, "brand": i.brand, "item": i.item_kpi,
                                        "weight_penetration": str(i.weight_pen),
                                        "sro": str(i.sro), "penetration": str(i.pen),
                                        "quantity": str(i.quantity), "amount_euro": str(i.amount_euro),"weighted_sro":str(i.weight_sro)}])
        with open(filename, "w", newline="", encoding='utf-16') as file:
            columns = ["year","month","brand","item","weight_penetration","sro","penetration","quantity","amount_euro","weighted_sro"]
            writer = csv.DictWriter(file, fieldnames=columns)
            writer.writeheader()
            for item in final_list:
                writer.writerows(item)

    def read_item_2020_OTC(conn,year):
        with sqlite3.connect("tertiary_sales_database.db") as conn:
            cursor = conn.cursor()
            cursor.execute(
                f"SELECT DISTINCT items.item_kpi_report, tertiary_sales.Year, tertiary_sales.PeriodName, items.brand, tertiary_sales.WeightPenetration,  tertiary_sales.SRO, tertiary_sales.Penetration, tertiary_sales.Quantity, tertiary_sales.Volume, tertiary_sales.WeightSRO from items join tertiary_sales on tertiary_sales.Fullmedicationname = items.item_proxima WHERE tertiary_sales.Year = {year} and items.sales_method = 'OTC'")
            results = cursor.fetchall()
            tertiary_list_otc = []
            for i in results:
                y_1 = str(i[4]).replace(',', '.')
                y_2 = str(i[5]).replace(',', '.')
                y_3 = str(i[6]).replace(',', '.')
                y_4 = str(i[7]).replace(',', '.')
                y_5 = str(i[8]).replace(',', '.')
                y_6 = str(i[9]).replace(',', '.')
                z = Tertiary_download_structure(i[0], i[1], i[2], i[3], y_1, y_2, y_3, y_4, y_5,y_6)

                tertiary_list_otc.append(z)
        return tertiary_list_otc

    def save_items_otc_to_csv_2020(self, filename,list_months,year):
        with sqlite3.connect("tertiary_sales_database.db") as conn:
            cursor = conn.cursor()
            cursor.execute(
                f"SELECT DISTINCT items.item_kpi_report, tertiary_sales.Year, tertiary_sales.PeriodName, items.brand, tertiary_sales.WeightPenetration,  tertiary_sales.SRO, tertiary_sales.Penetration, tertiary_sales.Quantity, tertiary_sales.Volume, tertiary_sales.WeightSRO from items join tertiary_sales on tertiary_sales.Fullmedicationname = items.item_proxima WHERE tertiary_sales.Year = {year} and items.sales_method = 'OTC'")
            results = cursor.fetchall()
            tertiary_list = []
            for i in results:
                y_1 = str(i[4]).replace(',', '.')
                y_2 = str(i[5]).replace(',', '.')
                y_3 = str(i[6]).replace(',', '.')
                y_4 = str(i[7]).replace(',', '.')
                y_5 = str(i[8]).replace(',', '.')
                y_6 = str(i[9]).replace(',', '.')
                z = Tertiary_download_structure(i[0], i[1], i[2], i[3], y_1, y_2, y_3, y_4, y_5,y_6)
                tertiary_list.append(z)
        final_list = []
        for month in list_months:
            for i in tertiary_list:
                if month == i.month:
                    final_list.append([{"year": str(i.year), "month": i.month, "brand": i.brand, "item": i.item_kpi,
                                        "weight_penetration": str(i.weight_pen),
                                        "sro": str(i.sro), "penetration": str(i.pen),
                                        "quantity": str(i.quantity), "amount_euro": str(i.amount_euro),"weighted_sro":str(i.weight_sro)}])
        with open(filename, "w", newline="", encoding='utf-16') as file:
            columns = ["year","month","brand","item","weight_penetration","sro","penetration","quantity","amount_euro","weighted_sro"]
            writer = csv.DictWriter(file, fieldnames=columns)
            writer.writeheader()
            for item in final_list:
                writer.writerows(item)
    def read_item_2020_OTC_with_commas(conn,year):
        with sqlite3.connect("tertiary_sales_database.db") as conn:
            cursor = conn.cursor()
            cursor.execute(
                f"SELECT DISTINCT items.item_kpi_report, tertiary_sales.Year, tertiary_sales.PeriodName, items.brand, tertiary_sales.WeightPenetration,  tertiary_sales.SRO, tertiary_sales.Penetration, tertiary_sales.Quantity, tertiary_sales.Volume, tertiary_sales.WeightSRO from items join tertiary_sales on tertiary_sales.Fullmedicationname = items.item_proxima WHERE tertiary_sales.Year = {year} and items.sales_method = 'OTC'")
            results = cursor.fetchall()
            tertiary_list_otc = []
            for i in results:
                y_1 = str(i[4])
                y_2 = str(i[5])
                y_3 = str(i[6])
                y_4 = str(i[7])
                y_5 = str(i[8])
                y_6 = str(i[9])
                z = Tertiary_download_structure(i[0], i[1], i[2], i[3], y_1, y_2, y_3, y_4, y_5,y_6)

                tertiary_list_otc.append(z)
        return tertiary_list_otc

    def save_items_otc_to_csv_2020_with_commas(self, filename,list_months,year):
        with sqlite3.connect("tertiary_sales_database.db") as conn:
            cursor = conn.cursor()
            cursor.execute(
                f"SELECT DISTINCT items.item_kpi_report, tertiary_sales.Year, tertiary_sales.PeriodName, items.brand, tertiary_sales.WeightPenetration,  tertiary_sales.SRO, tertiary_sales.Penetration, tertiary_sales.Quantity, tertiary_sales.Volume, tertiary_sales.WeightSRO from items join tertiary_sales on tertiary_sales.Fullmedicationname = items.item_proxima WHERE tertiary_sales.Year = {year} and items.sales_method = 'OTC'")
            results = cursor.fetchall()
            tertiary_list = []
            for i in results:
                y_1 = str(i[4])
                y_2 = str(i[5])
                y_3 = str(i[6])
                y_4 = str(i[7])
                y_5 = str(i[8])
                y_6 = str(i[9])
                z = Tertiary_download_structure(i[0], i[1], i[2], i[3], y_1, y_2, y_3, y_4, y_5,y_6)
                tertiary_list.append(z)
        final_list = []
        for month in list_months:
            for i in tertiary_list:
                if month == i.month:
                    final_list.append([{"year": str(i.year), "month": i.month, "brand": i.brand, "item": i.item_kpi,
                                        "weight_penetration": str(i.weight_pen),
                                        "sro": str(i.sro), "penetration": str(i.pen),
                                        "quantity": str(i.quantity), "amount_euro": str(i.amount_euro),"weighted_sro":str(i.weight_sro)}])
        with open(filename, "w", newline="", encoding='utf-16') as file:
            columns = ["year","month","brand","item","weight_penetration","sro","penetration","quantity","amount_euro","weighted_sro"]
            writer = csv.DictWriter(file, fieldnames=columns)
            writer.writeheader()
            for item in final_list:
                writer.writerows(item)

    def read_item_2020_RX_with_commas(conn,year):
        with sqlite3.connect("tertiary_sales_database.db") as conn:
            cursor = conn.cursor()
            cursor.execute(
                f"SELECT DISTINCT items.item_kpi_report, tertiary_sales.Year, tertiary_sales.PeriodName, items.brand, tertiary_sales.WeightPenetration,  tertiary_sales.SRO, tertiary_sales.Penetration, tertiary_sales.Quantity, tertiary_sales.Volume, tertiary_sales.WeightSRO from items join tertiary_sales on tertiary_sales.Fullmedicationname = items.item_proxima WHERE tertiary_sales.Year = {year} and items.sales_method = 'RX'")
            results = cursor.fetchall()
            tertiary_list_otc = []
            for i in results:
                y_1 = str(i[4])
                y_2 = str(i[5])
                y_3 = str(i[6])
                y_4 = str(i[7])
                y_5 = str(i[8])
                y_6 = str(i[9])
                z = Tertiary_download_structure(i[0], i[1], i[2], i[3], y_1, y_2, y_3, y_4, y_5,y_6)

                tertiary_list_otc.append(z)
        return tertiary_list_otc

    def save_items_RX_to_csv_2020_with_commas(self, filename,list_months,year):
        with sqlite3.connect("tertiary_sales_database.db") as conn:
            cursor = conn.cursor()
            cursor.execute(
                f"SELECT DISTINCT items.item_kpi_report, tertiary_sales.Year, tertiary_sales.PeriodName, items.brand, tertiary_sales.WeightPenetration,  tertiary_sales.SRO, tertiary_sales.Penetration, tertiary_sales.Quantity, tertiary_sales.Volume, tertiary_sales.WeightSRO from items join tertiary_sales on tertiary_sales.Fullmedicationname = items.item_proxima WHERE tertiary_sales.Year = {year} and items.sales_method = 'RX'")
            results = cursor.fetchall()
            tertiary_list = []
            for i in results:
                y_1 = str(i[4])
                y_2 = str(i[5])
                y_3 = str(i[6])
                y_4 = str(i[7])
                y_5 = str(i[8])
                y_6 = str(i[9])
                z = Tertiary_download_structure(i[0], i[1], i[2], i[3], y_1, y_2, y_3, y_4, y_5,y_6)

                tertiary_list.append(z)
        final_list = []
        for month in list_months:
            for i in tertiary_list:
                if month == i.month:
                    final_list.append([{"year": str(i.year), "month": i.month, "brand": i.brand, "item": i.item_kpi,
                                        "weight_penetration": str(i.weight_pen),
                                        "sro": str(i.sro), "penetration": str(i.pen),
                                        "quantity": str(i.quantity), "amount_euro": str(i.amount_euro),"weighted_sro":str(i.weight_sro)}])
        with open(filename, "w", newline="", encoding='utf-16') as file:
            columns = ["year","month","brand","item","weight_penetration","sro","penetration","quantity","amount_euro","weighted_sro"]
            writer = csv.DictWriter(file, fieldnames=columns)
            writer.writeheader()
            for item in final_list:
                writer.writerows(item)
    def read_item_2020_RX(conn,year):
        with sqlite3.connect("tertiary_sales_database.db") as conn:
            cursor = conn.cursor()
            cursor.execute(
                f"SELECT DISTINCT items.item_kpi_report, tertiary_sales.Year, tertiary_sales.PeriodName, items.brand, tertiary_sales.WeightPenetration,  tertiary_sales.SRO, tertiary_sales.Penetration, tertiary_sales.Quantity, tertiary_sales.Volume, tertiary_sales.WeightSRO from items join tertiary_sales on tertiary_sales.Fullmedicationname = items.item_proxima WHERE tertiary_sales.Year = {year} and items.sales_method = 'RX'")
            results = cursor.fetchall()
            tertiary_list_otc = []
            for i in results:
                y_1 = str(i[4]).replace(',', '.')
                y_2 = str(i[5]).replace(',', '.')
                y_3 = str(i[6]).replace(',', '.')
                y_4 = str(i[7]).replace(',', '.')
                y_5 = str(i[8]).replace(',', '.')
                y_6 = str(i[9]).replace(',', '.')
                z = Tertiary_download_structure(i[0], i[1], i[2], i[3], y_1, y_2, y_3, y_4, y_5,y_6)

                tertiary_list_otc.append(z)
        return tertiary_list_otc

    def save_items_RX_to_csv_2020(self, filename,list_months,year):
        with sqlite3.connect("tertiary_sales_database.db") as conn:
            cursor = conn.cursor()
            cursor.execute(
                f"SELECT DISTINCT items.item_kpi_report, tertiary_sales.Year, tertiary_sales.PeriodName, items.brand, tertiary_sales.WeightPenetration,  tertiary_sales.SRO, tertiary_sales.Penetration, tertiary_sales.Quantity, tertiary_sales.Volume, tertiary_sales.WeightSRO from items join tertiary_sales on tertiary_sales.Fullmedicationname = items.item_proxima WHERE tertiary_sales.Year = {year} and items.sales_method = 'RX'")
            results = cursor.fetchall()
            tertiary_list = []
            for i in results:
                y_1 = str(i[4]).replace(',', '.')
                y_2 = str(i[5]).replace(',', '.')
                y_3 = str(i[6]).replace(',', '.')
                y_4 = str(i[7]).replace(',', '.')
                y_5 = str(i[8]).replace(',', '.')
                y_6 = str(i[9]).replace(',', '.')
                z = Tertiary_download_structure(i[0], i[1], i[2], i[3], y_1, y_2, y_3, y_4, y_5,y_6)
                tertiary_list.append(z)
        final_list = []
        for month in list_months:
            for i in tertiary_list:
                if month == i.month:
                    final_list.append([{"year": str(i.year), "month": i.month, "brand": i.brand, "item": i.item_kpi,
                                        "weight_penetration": str(i.weight_pen),
                                        "sro": str(i.sro), "penetration": str(i.pen),
                                        "quantity": str(i.quantity), "amount_euro": str(i.amount_euro),"weighted_sro":str(i.weight_sro)}])
        with open(filename, "w", newline="", encoding='utf-16') as file:
            columns = ["year","month","brand","item","weight_penetration","sro","penetration","quantity","amount_euro","weighted_sro"]
            writer = csv.DictWriter(file, fieldnames=columns)
            writer.writeheader()
            for item in final_list:
                writer.writerows(item)

    def test_secondary_2021(self,month_ru):
        with sqlite3.connect("tertiary_sales_database.db") as conn:
            cursor = conn.cursor()
            cursor.execute(
                f"select secondary_sales_2021.item_quadra, secondary_sales_2021.cip_fact from secondary_sales_2021 where secondary_sales_2021.month_ru = '{month_ru}'")
            results = cursor.fetchall()
            secondary_list = []
            for i in results:
                y_1 = str(i[0]).replace(',', '.')
                y_2 = str(i[1]).replace(',', '.')
                z = Secondary_total_2021(y_1, y_2)
                secondary_list.append(z)
            sum = 0
            for i in secondary_list:
                sum += float(i.sales_euro)
            return sum
    def test_secondary_2020(self,month_ru):
        with sqlite3.connect("tertiary_sales_database.db") as conn:
            cursor = conn.cursor()
            cursor.execute(
                f"select secondary_sales_2020.item_quadra, secondary_sales_2020.cip_fact from secondary_sales_2020 where secondary_sales_2020.month_ru = '{month_ru}'")
            results = cursor.fetchall()
            secondary_list = []
            for i in results:
                y_1 = str(i[0]).replace(',', '.')
                y_2 = str(i[1]).replace(',', '.')
                z = Secondary_total_2021(y_1, y_2)
                secondary_list.append(z)
            sum = 0
            for i in secondary_list:
                sum += float(i.sales_euro)
            return sum

