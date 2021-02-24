import csv
import sqlite3
import pandas as pd
from pandas import Series
import xlsxwriter

import jaydebeapi
import pytest
import xlwt
from pandas.tests.io.excel.test_openpyxl import openpyxl


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

class Kam_plan_download_structure:
    def __init__(self,item_quadra, code_sf,month_local, quarter,plan_packs, plan_euro):
        self.plan_euro = plan_euro
        self.plan_packs = plan_packs
        self.quarter = quarter
        self.month_local = month_local
        self.code_sf = code_sf
        self.item_quadra = item_quadra


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
                 etalon_code_okpo,delivery_date,position_code,office_head_organization,head_office_okpo,quarter_year,half_year,
                 annual_sales_category,med_representative_name,kam_name,week,territory_name,brik_name,sale_in_quantity):
        self.position_code = position_code
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

class Quadra_from_xlxs_629():
    def __init__(self, year, month, ff_region, country_region, city_town, organization_name,
                 organization_adress, product_group,
                 product_code, item_quadra, organization_etalon_id, organization_etalon_name,
                 distributor_etalon_name,
                 distributor_name, distributor_okpo, sales_euro, promotion, organization_type,
                 organization_status,
                 etalon_code_okpo, delivery_date,position_code, office_head_organization, head_office_okpo,
                 quarter_year, half_year,
                 annual_sales_category, med_representative_name, kam_name, week, territory_name,
                 brik_name, sales_packs):
        self.position_code = position_code
        self.product_group = product_group
        self.sales_packs = sales_packs
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
        self.sales_euro = sales_euro
        self.distributor_okpo = distributor_okpo
        self.distributor_name = distributor_name
        self.distributor_etalon_name = distributor_etalon_name
        self.organization_etalon_name = organization_etalon_name
        self.organization_etalon_id = organization_etalon_id
        self.item_quadra = item_quadra
        self.product_code = product_code
        self.organization_adress = organization_adress
        self.organization_name = organization_name
        self.city_town = city_town
        self.country_region = country_region
        self.ff_region = ff_region
        self.month = month
        self.year = year
    def __str__(self):
        return f'{self.year}, {self.month}, {self.ff_region}, ' \
               f'{self.country_region}, {self.city_town}, {self.organization_name},' \
               f'{self.organization_adress}, {self.product_group},{self.product_code}, ' \
               f'{self.item_quadra}, {self.organization_etalon_id}, {self.organization_etalon_name},' \
               f'{self.distributor_etalon_name},{self.distributor_name},{self.distributor_okpo},' \
               f'{self.sales_euro},{self.promotion},{self.organization_type},{self.organization_status},' \
               f'{self.etalon_code_okpo}, {self.delivery_date}, {self.position_code}, {self.office_head_organization}, ' \
               f'{self.head_office_okpo},{self.quarter_year}, {self.half_year},{self.annual_sales_category},' \
               f'{self.med_representative_name},{self.kam_name},{self.week},{self.territory_name},' \
               f'{self.brik_name}, {self.sales_packs}'
    def __repr__(self):
        return f'{self.year}, {self.month}, {self.ff_region}, {self.country_region}, {self.city_town}, {self.organization_name},{self.organization_adress}, {self.product_group},{self.product_code}, {self.item_quadra}, {self.organization_etalon_id}, {self.organization_etalon_name},{self.distributor_etalon_name},{self.distributor_name},{self.distributor_okpo},{self.sales_euro},{self.promotion},{self.organization_type},{self.organization_status},{self.etalon_code_okpo}, {self.delivery_date}, {self.position_code},{self.office_head_organization}, {self.head_office_okpo},{self.quarter_year}, {self.half_year},{self.annual_sales_category},{self.med_representative_name},{self.kam_name},{self.week},{self.territory_name},{self.brik_name}, {self.sales_packs}'

class CBase_2021_quadra_workout:
    def classify_base_2021(self,base_2021):
        base_2021_classifyed = []
        for i in base_2021:
            sale_in_quantity = i[32]
            brik_name = i[31]
            territory_name = i[30]
            week = i[29]
            kam_name = i[28]
            med_representative_name = i[27]
            annual_sales_category = i[26]
            half_year = i[25]
            quarter_year = i[24]
            head_office_okpo = i[23]
            office_head_organization = i[22]
            delivery_date = i[20]
            position_code = i[21]
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
                     etalon_code_okpo,delivery_date,position_code,office_head_organization,head_office_okpo,quarter_year,half_year,
                     annual_sales_category,med_representative_name,kam_name,week,territory_name,brik_name,sale_in_quantity)
            base_2021_classifyed.append(entry)
        return base_2021_classifyed
    def classify_base_2021_from_xlxs(self, item):
        base_2021_classified_ = []
        year = str(item[0])
        month = str(item[1])
        ff_region = str(item[3])
        country_region = str(item[2])
        city_town = str(item[4])
        organization_name = str(item[5])
        organization_adress = str(item[6])
        product_group = str(item[7])
        product_code = str(item[8])
        item_quadra = str(item[9])
        organization_etalon_id = str(item[10])
        organization_etalon_name = str(item[11])
        distributor_etalon_name = str(item[12])
        distributor_name = str(item[13])
        distributor_okpo = str(item[14])
        sales_euro = str(item[15]).replace(',','.')
        promotion = str(item[16])
        organization_type = str(item[17])
        organization_status = str(item[18])
        etalon_code_okpo = str(item[19])
        delivery_date = str(item[20])
        position_code = str(item[21])
        office_head_organization = str(item[22])
        head_office_okpo = str(item[23])
        quarter_year = str(item[24])
        half_year = str(item[25])
        annual_sales_category = str(item[26])
        med_representative_name = str(item[27])
        kam_name = str(item[28])
        week = str(item[29])
        territory_name = str(item[30])
        brik_name = str(item[31])
        sales_packs = str(item[32]).replace(',','.')

        st = Quadra_from_xlxs_629(year, month, ff_region, country_region, city_town, organization_name,
               organization_adress, product_group,
               product_code, item_quadra, organization_etalon_id, organization_etalon_name,
               distributor_etalon_name,
               distributor_name, distributor_okpo, sales_euro, promotion, organization_type,
               organization_status,
               etalon_code_okpo, delivery_date, position_code, office_head_organization, head_office_okpo,
               quarter_year, half_year,
               annual_sales_category, med_representative_name,kam_name, week, territory_name,
               brik_name, sales_packs)

        base_2021_classified_.append(st)
        return base_2021_classified_


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
                           + " empl.bar_code AS position_code,"  # added position code
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
                           + " empl.bar_code AS position_code,"  # added position code
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

    def save_base_629_2021_to_xlsx(self):
        final_list = []

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
                       + " empl.bar_code AS position_code,"  # added position code
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

        sum_tot_euro = 0
        workbook = xlsxwriter.Workbook('C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\0.new_629_report_2021.xlsx')
        worksheet = workbook.add_worksheet()

        # Widen the first column to make the text clearer.
        #worksheet.set_column('A:A', 20)
        bold = workbook.add_format({'bold': True},)
        worksheet.write('A1', "Год", bold)
        worksheet.write('B1', "Месяц", bold)
        worksheet.write('C1', "Область", bold)
        worksheet.write('D1', "Регион", bold)
        worksheet.write('E1', "Населенный пункт", bold)
        worksheet.write('F1', "Организация", bold)
        worksheet.write('G1', "Почтовый адрес (организация)", bold)
        worksheet.write('H1', "Группа товара", bold)
        worksheet.write('I1', "product_code", bold)
        worksheet.write('J1', 'Товар', bold)
        worksheet.write('K1', 'Организация(эталонId)', bold)
        worksheet.write('L1', 'Организация(эталон)', bold)
        worksheet.write('M1', 'Дистрибьютор (эталон)', bold)
        worksheet.write('N1', 'Дистрибьютор', bold)
        worksheet.write('O1', 'ОКПО (дистрибъютор)', bold)
        worksheet.write('P1', 'Сумма (IN)', bold)
        worksheet.write('Q1', 'Группа товара 2', bold)
        worksheet.write('R1', 'Вид организации', bold)
        worksheet.write('S1', 'Тип организации', bold)
        worksheet.write('T1', 'ОКПО', bold)
        worksheet.write('U1', 'Дата отгрузки', bold)
        worksheet.write('V1', 'Код позиции', bold)
        worksheet.write('W1', 'Гл. Офис Сети', bold)
        worksheet.write('X1', 'Гл. Офис ОКПО', bold)
        worksheet.write('Y1', 'Квартал', bold)
        worksheet.write('Z1', 'Полугодие', bold)
        worksheet.write('AA1', 'Категория товарооборота', bold)
        worksheet.write('AB1', 'Сотрудник', bold)
        worksheet.write('AC1', 'КАМ', bold)
        worksheet.write('AD1', 'Неделя', bold)
        worksheet.write('AE1', 'Территория', bold)
        worksheet.write('AF1', 'Полигон', bold)
        worksheet.write('AG1', 'Количество (IN)', bold)


        columns = ['year', 'month', 'ff_region', 'country_region', 'city_town', 'organization_name',
                   'organization_adress', 'sales_method',
                   'product_code', 'item_quadra', 'organization_etalon_id', 'organization_etalon_name',
                   'distributor_etalon_name',
                   'distributor_name', 'distributor_okpo', 'sales_euro_', 'promotion', 'organization_type',
                   'organization_status',
                   'etalon_code_okpo', 'delivery_date', 'position_code','office_head_organization', 'head_office_okpo',
                   'quarter_year', 'half_year',
                   'annual_sales_category', 'med_representative_name', 'kam_name', 'week', 'territory_name',
                   'brik_name', 'sale_in_quantity']

        list_base_2021 = []
        row_index = 1

        for item in base_2021_classifyed_:
            z0 = str(item.year).replace(",", "")
            y1 = str(item.sales_euro_).replace(".", ",")
            y2 = str(item.sale_in_quantity).replace(".", ",")
            item_ = [[str(z0),
                     str(item.month),
                     str(item.ff_region),
                     str(item.country_region),
                     str(item.city_town),
                     str(item.organization_name),
                     str(item.organization_adress),
                     str(item.sales_method),
                     str(item.product_code),
                     str(item.item_quadra),
                     str(item.organization_etalon_id),
                     str(item.organization_etalon_name),
                     str(item.distributor_etalon_name),
                     str(item.distributor_name),
                     str(item.distributor_okpo),
                     str(y1),
                     str(item.promotion),
                     str(item.organization_type),
                     str(item.organization_status),
                     str(item.etalon_code_okpo),
                     str(item.delivery_date),
                     str(item.position_code),
                     str(item.office_head_organization),
                     str(item.head_office_okpo),
                     str(item.quarter_year),
                     str(item.half_year),
                     str(item.annual_sales_category),
                     str(item.med_representative_name),
                     str(item.kam_name),
                     str(item.week),
                     str(item.territory_name),
                     str(item.brik_name),
                     str(y2)]]

            list_base_2021.append(item_)
            worksheet.write(int(row_index), int(0),str(z0))
            worksheet.write(int(row_index), int(1),str(item.month))
            worksheet.write(int(row_index), int(2),str(item.ff_region))
            worksheet.write(int(row_index), int(3),str(item.country_region))
            worksheet.write(int(row_index), int(4),str(item.city_town))
            worksheet.write(int(row_index), int(5),str(item.organization_name))
            worksheet.write(int(row_index), int(6),str(item.organization_adress))
            worksheet.write(int(row_index), int(7),str(item.sales_method))
            worksheet.write(int(row_index), int(8),str(item.product_code))
            worksheet.write(int(row_index), int(9),str(item.item_quadra))
            worksheet.write(int(row_index), int(10),str(item.organization_etalon_id))
            worksheet.write(int(row_index), int(11),str(item.organization_etalon_name))
            worksheet.write(int(row_index), int(12),str(item.distributor_etalon_name))
            worksheet.write(int(row_index), int(13),str(item.distributor_name))
            worksheet.write(int(row_index), int(14),str(item.distributor_okpo))
            worksheet.write(int(row_index), int(15),str(y1))
            worksheet.write(int(row_index), int(16),str(item.promotion))
            worksheet.write(int(row_index), int(17),str(item.organization_type))
            worksheet.write(int(row_index), int(18),str(item.organization_status))
            worksheet.write(int(row_index), int(19),str(item.etalon_code_okpo))
            worksheet.write(int(row_index), int(20),str(item.delivery_date))
            worksheet.write(int(row_index), int(21),str(item.position_code))
            worksheet.write(int(row_index), int(22),str(item.office_head_organization))
            worksheet.write(int(row_index), int(23),str(item.head_office_okpo))
            worksheet.write(int(row_index), int(24),str(item.quarter_year))
            worksheet.write(int(row_index), int(25),str(item.half_year))
            worksheet.write(int(row_index), int(26),str(item.annual_sales_category))
            worksheet.write(int(row_index), int(27),str(item.med_representative_name))
            worksheet.write(int(row_index), int(28),str(item.kam_name))
            worksheet.write(int(row_index), int(29),str(item.week))
            worksheet.write(int(row_index), int(30),str(item.territory_name))
            worksheet.write(int(row_index), int(31),str(item.brik_name))
            worksheet.write(int(row_index), int(32),str(y2))

            row_index +=1

        workbook.close()

    def get_secondary_2021_by_month(self):
        path = "C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\0.new_629_report_2021.xlsx"
        wb_obj = openpyxl.load_workbook(path)
        sheet_obj = wb_obj.active
        rows_count = str(sheet_obj.calculate_dimension()).rsplit(':')
        rows_count = int(str(rows_count[1])[2:])
        print(rows_count)
        string = []
        classified_base_2021 = []
        for row in range(1, rows_count + 1):
            str_ = []
            for col in range(1, 34):
                cell_obj = sheet_obj.cell(row=row, column=col)
                str_.append(cell_obj.value)
            string.append(str_)
        for i in string:
            x = CBase_2021_quadra_workout()
            string_class = x.classify_base_2021_from_xlxs(i)
            classified_base_2021.append(string_class)
        return classified_base_2021

    def get_secondary_2020_by_month(self):
        path = "C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\0.new_629_report_2020.xlsx"
        wb_obj = openpyxl.load_workbook(path)
        sheet_obj = wb_obj.active
        rows_count = str(sheet_obj.calculate_dimension()).rsplit(':')
        rows_count = int(str(rows_count[1])[2:])
        string = []
        classified_base_2020 = []
        for row in range(1, rows_count + 1):
            str_ = []
            for col in range(1, 33):
                cell_obj = sheet_obj.cell(row=row, column=col)
                str_.append(cell_obj.value)
            string.append(str_)
        for i in string:
            x = CBase_2021_quadra_workout()
            string_class = x.classify_base_2021_from_xlxs(i)
            classified_base_2020.append(string_class)
        return classified_base_2020

    #def get_629_from_xlxs(self):







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



    def print_actual_MTD_sales(base_raw,year,month_local):
        report_test = CTest_SAles_report_creation()
        report = report_test.test_rep(base_raw)
        sales_euro = 0
        week = ''
        distr = ''
        item_quadra = ''
        quantity = 0
        for i in report:
            if i.year == year and i.month == month_local:
                sales_euro += i.sec_euro
                quantity += i.sec_quantity
        print(f"Год: {year}\nМесяц: {month_local}\nОбщие продажи в упаковках:", '{0:,}'.format(quantity.__round__(2)).replace(",", " "), "packs\nОбщие продажи в евро:", '{0:,}'.format(sales_euro.__round__(2)).replace(",", " "), 'euro')



class CEXtract_database_tertiary:
    def read_item(conn, year):
        tertiary_list = []
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
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
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
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
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
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
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
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
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
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
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
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
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
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
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
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
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
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
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
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
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
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
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
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
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
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
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
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
class Kam_plans:
    def read_kam_plan(conn):
        plan_list = []
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
            cursor = conn.cursor()
            cursor.execute(f"select distinct kam_plan.item_quadra, kam_plan.code, kam_plan.month,  ymm.Квартал, kam_plan.plan_packs, kam_plan.plan_euro from kam_plan join ymm on ymm.Месяц = kam_plan.month")
            results = cursor.fetchall()
            for i in results:
                y_1 = str(i[0])
                y_2 = str(i[1])
                y_3 = str(i[2])
                y_4 = str(i[3])
                y_5 = str(i[4])
                y_6 = str(i[5]).replace(',', '.')
                z = Kam_plan_download_structure(y_1, y_2, y_3,y_4, y_5, y_6)
                plan_list.append(z)
        return plan_list