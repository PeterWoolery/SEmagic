# SEmagic
Programatically interact with Smart Etailing

## Requirements:
- Selemium - Interacting with smart etailing via web debugging framework
- Chrome or Chromium browser - Likely to be most supported browser for Smart Etailing moving forward

## webdriver
This is the main method of interacting with SmartEtailin in the same way a human would. This should only be used in situations where a documented API does not exist... which is most situations.

**start** - creates a webdriver instance in headless mode. Can be specified as non-headless, specific window size, and to use a proxy

**login** - login to SE using credentials provided in config.yaml This ia a required step before most other fuctions

**new_orders** - opens orders page and checks all the boxes on the first page where the order status is "New Order Pending". This stages these orders to populate the official XML order export API. NOTE: unlike the official export method, this ignores every aspect other that the order status, so it will include orders prior to requesting fulfillment.

**capture_payments** - opens an order in SE and clicks capture payment for those that are made via Authorize.net

**request_fulfillment** - requests fulfillment for items on an order for QBP and Trek. HLC could be built in later, but It may be better to use their official API unless you really want feedback from SE's integration. HLC's API is well documented and reliable.

**change_shipping** - simple function to change the shipping type from Home Delivery to Ground Shipping to allow for QBP and Trek fulfillment of these orders. If you submit a fulfillment request when home delivery is set, SE will get a notification from Q that they refused the order.

**get_discount_ID** - pulls the discount description from an order's page

**update_shipping_cost** - In Development - method to change the "Additional Shipping" field of an item. This can be used to set up additional logic to your shipping policies. By default, SE allows dimensional weight, top level categories, and price thresholds. So if you want different shipping costs for different prices or types of bikes, you have to do it this way.

**create_unique_discount** - creates a unique discount code that is able to be used once on any item or order.

**delete_old_discounts** - removes "EXPIRED" discounts from commerce-discounts page. When using many unique discounts, that page can get unweily very fast. Currently, this is set to start at page 11 and run until no more are left.

**get_product_map** - pulls SKU/UPC/MPN product mapping file. This can be usefull for identifying which products are in your catalog and mapping correctly, or can be used to get itemIDs for modifying their pages.

**update_item_notes** - In Development - updates the notes of an item within an order.

**update_ordernotes** - Appends additional information to the bottom of the order notes.

**get_config** - parses config.yaml and returns a dictionary.