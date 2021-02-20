import sys
import jaydebeapi


def main():
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
+" month.name AS month,"
+" region.name AS ff_region,"
+" province.name AS country_region,"
+" town.name AS city_town,"
+" organization.name AS organization_name,"
+" organization.postal_address AS organization_adress,"
+" category.name AS sales_method,"
+" product.code AS product_code,"
+" product.name AS item_quadra,"
+" orgSalesByProduct.organization_etalon_id AS organization_etalon_id,"
+" orgSalesByProduct.organization_etalon_name AS organization_etalon_name,"
+" orgSalesByProduct.distributor_etalon_name AS distributor_etalon_name,"
+" distributor.name AS distributor_name,"
+" orgSalesByProduct.distributor_okpo AS distributor_okpo,"
+" productPriceDynamics.price_fact * ROUND(orgSalesByProduct.quantity, 2) AS sales_euro_,"
+" productCategory2.name AS promotion,"
+" organizationType.name AS organization_type,"
+" organizationStatus.name AS organization_status,"
+" orgSalesByProduct.etalon_code_okpo AS etalon_code_okpo,"
+" orgSalesByProduct.delivery_date AS delivery_date,"
+" 	   "
+" ("
+" SELECT MAX(orgOf.name+' '+sType.name+' '+orgOf.street_name+', '+orgOf.street_number)"
+" FROM organization AS orgOf"
+"     LEFT JOIN street_type sType ON sType.id = orgOf.street_type_id"
+" WHERE orgOf.etalon_id = orgSalesByProduct.head_organization_etalon_id"
+" ) AS office_head_organization,"
+" CASE"
+"     WHEN orgSalesByProduct.head_office_okpo IS NOT NULL"
+"     THEN orgSalesByProduct.head_office_okpo"
+"     ELSE ''"
+" END AS head_office_okpo,"
+" CASE"
+"     WHEN month.id = 1"
+"         OR month.id = 2"
+"         OR month.id = 3"
+"     THEN '1-й квартал'"
+"     ELSE CASE"
+"             WHEN month.id = 4"
+"                     OR month.id = 5"
+"                     OR month.id = 6"
+"             THEN '2-й квартал'"
+"             ELSE CASE"
+"                         WHEN month.id = 7"
+"                             OR month.id = 8"
+"                             OR month.id = 9"
+"                         THEN '3-й квартал'"
+"                         ELSE CASE"
+"                                 WHEN month.id = 10"
+"                                     OR month.id = 11"
+"                                     OR month.id = 12"
+"                                 THEN '4-й квартал'"
+"                                 ELSE ''"
+"                             END"
+"                     END"
+"         END"
+" END AS quarter_year,"
+" "
+" CASE"
+"     WHEN month.id IN(1, 2, 3, 4, 5, 6)"
+"     THEN '1-е полугодие'"
+"     ELSE '2-е полугодие'"
+" END AS half_year,"
+" "
+" ("
+" SELECT sbm.name"
+" FROM sales_by_month_category AS sbm"
+" WHERE sbm.pos_status = 0"
+"     AND sbm.id ="
+" ("
+" SELECT orgfet.sales_by_month_category_id"
+" FROM organization_from_etalon AS orgfet"
+" WHERE orgfet.pos_status = 0"
+"         AND orgfet.etalon_id = organization.etalon_id"
+" )"
+" ) AS annual_sales_category,"
+" "
+" "
+" "
+" CASE"
+"     WHEN empl.id IS NOT NULL"
+"         AND emplSG.role_id = 3"
+"         AND empl.status_id = 1"
+"     THEN empl.last_name+' '+empl.first_name"
+"     ELSE ''"
+" END AS med_representative_name,"
+" CASE"
+"     WHEN kam.id IS NOT NULL"
+"     THEN kam.last_name+' '+kam.first_name"
+"     ELSE ''"
+" END AS kam_name,"
+" DATEPART(ISO_WEEK, orgSalesByProduct.delivery_date) AS week,"
+" territory.name AS territory_name,"
+" brick.name AS brik_name,"
+" ROUND(orgSalesByProduct.quantity, 2) AS sale_in_quantity"
+" "
+" "
+" "
+" FROM organization_sales_by_product AS orgSalesByProduct"
+" LEFT JOIN month_classifier month ON orgSalesByProduct.month_id = month.id"
+" LEFT JOIN organization organization ON organization.id = orgSalesByProduct.organization_id"
+" LEFT JOIN registration_type registration ON organization.registration_type_id = registration.id"
+" LEFT JOIN rayon rayon ON organization.rayon_id = rayon.id"
+" LEFT JOIN town town ON town.id = orgSalesByProduct.town_id"
+" LEFT JOIN oblast province ON province.id = town.oblast_id"
+" LEFT JOIN region region ON region.id = province.region_id"
+" LEFT JOIN region region2 ON region2.id = province.region2_id"
+" LEFT JOIN region region3 ON region3.id = province.region3_id"
+" LEFT JOIN region region4 ON region4.id = province.region4_id"
+" LEFT JOIN oblast_rayon admRayon ON admRayon.id = town.oblast_rayon_id"
+" LEFT JOIN organization_type organizationType ON organizationType.id = organization.organization_type_id"
+" LEFT JOIN organization_status organizationStatus ON organizationStatus.id = organizationType.organization_status_id"
+" LEFT JOIN product product ON product.id = orgSalesByProduct.product_id"
+" LEFT JOIN product_series series ON series.id = product.product_series_id"
+" LEFT JOIN product_series_group psGroup ON psGroup.id = series.product_series_group_id"
+" LEFT JOIN product_category category ON category.id = product.product_category_id"
+" LEFT JOIN product_category2 category2 ON category2.id = product.product_category2_id"
+" LEFT JOIN product_category_atc productCategoryAtc ON productCategoryAtc.id = product.product_category_atc_id"
+" LEFT JOIN employee empl ON empl.id = orgSalesByProduct.employee_id"
+" LEFT JOIN security_group emplSG ON emplSG.id = empl.security_group_id"
+" LEFT JOIN organization_drugstore_office odo ON odo.organization_id = orgSalesByProduct.head_organization_id"
+"                                             AND odo.pos_status = 0"
+"                                             AND odo.active = 1"
+" LEFT JOIN employee kam ON kam.id = odo.manager_id"
+" LEFT JOIN organization_sales_by_product_op osbop ON osbop.organization_etalon_id = organization.etalon_id"
+"                                                     AND osbop.product_id = product.id"
+" LEFT JOIN territory_brick brick ON osbop.territory_brick_id = brick.id"
+" LEFT JOIN territory territory ON brick.territory_id = territory.id"
+" LEFT JOIN town headTown ON orgSalesByProduct.head_organization_town_id = headTown.id"
+" LEFT JOIN oblast headObl ON headTown.oblast_id = headObl.id"
+" LEFT JOIN organization distributor ON orgSalesByProduct.distributor_id = distributor.id"
+" LEFT JOIN sales_by_month_category salesByMonthCategory ON salesByMonthCategory.id = organization.sales_by_month_category_id"
+" LEFT JOIN category orgCat ON orgCat.id = organization.organization_category_id"
+" LEFT JOIN product_price_dynamics productPriceDynamics ON productPriceDynamics.pos_status = 0"
+"                                                         AND productPriceDynamics.month_id = orgSalesByProduct.month_id"
+"                                                         AND productPriceDynamics.year = orgSalesByProduct.year"
+"                                                         AND productPriceDynamics.product_id = orgSalesByProduct.product_id"
+" LEFT JOIN product_category2 productCategory2 ON productPriceDynamics.product_category2_id = productCategory2.id"
+" LEFT JOIN sales_type salesType ON organization.id = salesType.organization_id"
+"                                 AND salesType.year = orgSalesByProduct.year"
+"                                 AND salesType.month_id = month.id"
+"                                 AND salesType.pos_status = 0"
+" LEFT JOIN organization_specialization orgSpecialization ON organization.organization_specialization_id = orgSpecialization.id"
+" WHERE orgSalesByProduct.year = 2021")
        res = cursor.fetchall()

        print(str(res))


    except Exception as err:
        print(str(err))


if __name__ == "__main__":
    sys.exit(main())