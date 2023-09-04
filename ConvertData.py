import pandas as pd
import json

def read_tokopedia_xlsx(filepath: str, start_row: int):
    data = pd.read_excel(filepath, header=None)

    # Mengambil seluruh tabel data (mulai dari start_row)
    data = data.iloc[start_row-1:]

    # Mengambil baris pertama sebagai nama kolom
    data.columns = data.iloc[0]
    data = data[1:]

    # Mengonversi DataFrame ke dictionary
    data_dict = data.to_dict(orient='records')

    return data_dict

def tokopedia_to_order(data):
    orders = []
    for record in data:
        order = {
            "title": "Invoice",
            "invoice_number": record['Nomor Invoice'],
            "sender_name": "KremlinOfficial",
            "sender_phonenumber": "KremlinOfficial",
            "sender_address": "KremlinOfficial",
            "column_break_ndpbe": "---------------------",
            "date": record['Tanggal Pembayaran'],
            "receiver_name": record['Nama Penerima'],
            "receiver_phonenumber": record['No Telp Penerima'],
            "receiver_address": record['Alamat Pengiriman'],
            "marketplace": "Tokopedia",
            "courier": record['Nama Kurir'],
            "courier_service": record['Tipe Pengiriman (regular, same day, etc)'],
            "weight": "",
            "shipping_charge": record['Biaya Pengiriman Tunai (IDR)'],
            "airwaybill": record['No Resi / Kode Booking'],
            "is_cash": record['COD'],
            "order_item": [
                {
                    "sku": record['Nomor SKU'],
                    "name1": record['Nama Produk'],
                    "variant": record['Tipe Produk'],
                    "qty": record['Jumlah Produk Dibeli']
                }
            ],
            "note": record['Catatan Produk Pembeli'],
            "status": record['Status Terakhir']
        }

        orders.append(order)
        print(json.dumps(order, indent=2))  # Print JSON representation of each order

    return orders

def read_shopee_xlsx(filepath: str):
    data = pd.read_excel(filepath, header=None)
        
    # Mengambil baris pertama sebagai nama kolom
    data.columns = data.iloc[0]
    
    # Mengonversi DataFrame ke dictionary
    data_dict = data.to_dict(orient='records')
    
    return data_dict

def shopee_to_order(data):
    orders = []
    for record in data:
        order = {
            "title": "Invoice",
            "invoice_number": record['No. Pesanan'],
            "sender_name": "KremlinOfficial",
            "sender_phonenumber": "KremlinOfficial",
            "sender_address": "KremlinOfficial",
            "column_break_ndpbe": "---------------------",
            "date": record['Waktu Pembayaran Dilakukan'],
            "receiver_name": record['Nama Penerima'],
            "receiver_phonenumber": record['No. Telepon'],
            "receiver_address": record['Alamat Pengiriman'],
            "marketplace": "Shopee",
            "courier": record['Opsi Pengiriman'],
            "courier_service": "",
            "weight": record['Total Berat'],
            "shipping_charge": record['Perkiraan Ongkos Kirim'],
            "airwaybill": record['No. Resi'],
            "is_cash": "",
            "order_item": [
                {
                    "sku": record['SKU Induk'],
                    "name1": record['Nama Produk'],
                    "variant": record['Nama Variasi'],
                    "qty": record['Jumlah Produk di Pesan']
                }
            ],
            "note": record['Catatan dari Pembeli'],
            "status": record['Status Pesanan']
        }

        orders.append(order)
        print(json.dumps(order, indent=2))  # Print JSON representation of each order

    return orders

def read_marketplace_xlsx(filepath: str):
    data = pd.read_excel(filepath, header=None)
        
    # Mengambil baris pertama sebagai nama kolom
    data.columns = data.iloc[0]
    
    # Mengonversi DataFrame ke dictionary
    data_dict = data.to_dict(orient='records')
    
    return data_dict

def marketplace_to_order(data):
    orders = []
    for record in data:
        order = {
            "title": "Invoice",
            "invoice_number": record['Order ID'],
            "sender_name": "KremlinOfficial",
            "sender_phonenumber": "KremlinOfficial",
            "sender_address": "KremlinOfficial",
            "column_break_ndpbe": "---------------------",
            "date": record['Created Time'],
            "receiver_name": record['Recipient'],
            "receiver_phonenumber": record['Phone #'],
            "receiver_address": record['Detail Address'],
            "marketplace": "Marketplace",
            "courier": record['Shipping Provider Name'],
            "courier_service": record['Delivery Option'],
            "weight": record['Weight(kg)'],
            "shipping_charge": record['Perkiraan Ongkos Kirim'],
            "airwaybill": record['Tracking ID'],
            "is_cash": record['Payment Method'],
            "order_item": [
                {
                    "sku": record['SKU ID'],
                    "name1": record['Product Name'],
                    "variant": record['Variation'],
                    "qty": record['Quantity']
                }
            ],
            "note": "",
            "status": record['Order Status']
        }

        orders.append(order)
        print(json.dumps(order, indent=2))  

    return orders
    
def read_lazada_xlsx(filepath: str):
    data = pd.read_excel(filepath, header=None)
        
    data.columns = data.iloc[0]
    
    data_dict = data.to_dict(orient='records')
    
    return data_dict

def lazada_to_order(data):
    orders = []
    for record in data:
        order = {
            "title": "Invoice",
            "invoice_number": record['invoiceNumber'],
            "sender_name": "KremlinOfficial",
            "sender_phonenumber": "KremlinOfficial",
            "sender_address": "KremlinOfficial",
            "column_break_ndpbe": "---------------------",
            "date": record['createTime'],
            "receiver_name": record['customerName'],
            "receiver_phonenumber": record['shippingPhone'],
            "receiver_address": record['shippingAddress'],
            "marketplace": "Lazada",
            "courier": record['shippingProvider'],
            "courier_service":record['shippingProviderType'],
            "weight": "",
            "shipping_charge": record['shippingFee'],
            "airwaybill": record['trackingCode'],
            "is_cash": record['payMethod'],
            "order_item": [
                {
                    "sku": record['lazadaSku'],
                    "name1": record['itemName'],
                    "variant": record['variation'],
                    "qty": ""
                }
            ],
            "note": "",
            "status": record['status']
        }

        orders.append(order)
        print(json.dumps(order, indent=2))  

    return orders
    

if __name__ == '__main__':
    start_row = 5  
    data_tokped = read_tokopedia_xlsx("D:\\Bootcamp\\Projek Kremlin\\File Excel\\Tokopedia_Order_20230824-20230825.xlsx", start_row)
    data_shopee = read_shopee_xlsx("D:\\Bootcamp\\Projek Kremlin\\File Excel\\Order.toship.20230824_20230825.xlsx")
    data_lazada = read_lazada_xlsx("D:\\Bootcamp\\Projek Kremlin\\File Excel\\790233ab5e7931625786f55c6080d37d.xlsx")
    data_marketplace = read_marketplace_xlsx("D:\\Bootcamp\\Projek Kremlin\\File Excel\\Untuk Dikirim pesanan-2023-08-25-11_11.xlsx")
    # print(data_marketplace[0].keys())
    result_lazada = lazada_to_order(data_lazada)
    result_tokped = tokopedia_to_order(data_tokped)
    result_shopee = shopee_to_order(data_shopee)
    result_marketplace = marketplace_to_order(data_marketplace)
    