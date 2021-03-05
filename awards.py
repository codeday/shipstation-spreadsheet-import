import os
import gspread
from dotenv import load_dotenv
load_dotenv()

gc = gspread.service_account('service-account.json')
sheet = gc.open('Awards w2021')
worksheet = sheet.get_worksheet(0)
records = worksheet.get_all_records()

from shipstation.api import ShipStation, ShipStationOrder, ShipStationAddress, ShipStationWeight, ShipStationContainer
ss = ShipStation(key=os.getenv('API_KEY'), secret=os.getenv('API_SECRET'))
for record in records:
    if record['Status'] == 'awaiting shipment':
        ss_order = ShipStationOrder(order_number='w2021award')
        ss_order.set_status('awaiting_shipment')
        ss_weight = ShipStationWeight(units='ounces', value=4)
        ss_container = ShipStationContainer(
            units='inches',
            length=12.25,
            width=9.75,
            height=0.25
        )
        ss_container.set_weight(ss_weight)
        shipping_address = ShipStationAddress(
            name=record['Name'],
            street1=record['addrline1'],
            street2=record['addrline2'],
            city=record['City'],
            state=record['State'],
            postal_code=record['ZIP'],
            country=record['Country']
        )
        ss_order.set_billing_address(shipping_address)
        ss_order.set_shipping_address(shipping_address)
        ss_order.set_dimensions(ss_container)
        ss.add_order(ss_order)
        row = worksheet.find(str(record['discord ID'])).row
        worksheet.update_cell(row=row, col=12, value='in shipstation')

ss.submit_orders()