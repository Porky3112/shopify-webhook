from flask import Flask, request, jsonify
import os
from datetime import datetime

# Importa tu clase generadora de facturas (ajusta la ruta si está en otro archivo)
from generador_facturas import ShopifyInvoiceGenerator

app = Flask(__name__)

# Configuración de tu tienda Shopify (usa variables de entorno en producción)
config = {
    'shop_domain': 'cshop.co',
    'shopify_access_token': 'shpat_tu_token',
    'company_name': 'CSHOP SAS',
    'company_address': 'CRA 34 3-65',
    'company_phone': '+57 3158103812',
    'company_email': 'info@cshop.co',
    'company_tax_id': 'NIT: 901410087-8'
}

@app.route('/webhook', methods=['POST'])
def handle_webhook():
    try:
        data = request.json
        order_id = data['id']
        print(f"¡Nueva orden recibida! ID: {order_id}")

        # 1. Generar la factura
        generator = ShopifyInvoiceGenerator(config)
        invoice_path = generator.generate_invoice(
            order_id=str(order_id),
            save_local=True,  # Guardar localmente (o subir a la nube)
            upload_to_cloud=False
        )
        print(f"Factura generada en: {invoice_path}")

        # 2. Opcional: Enviar por email al cliente
        # enviar_email(data['customer']['email'], invoice_path)

        return jsonify({"status": "Factura generada", "path": invoice_path}), 200

    except Exception as e:
        print(f"Error al procesar la orden: {e}")
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=os.getenv("PORT", 5000))