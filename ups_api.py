import requests
import json
import os
import uuid
import logging
import datetime
from dotenv import load_dotenv

load_dotenv()

logging.basicConfig(
    filename='ups_api.log',
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

class UPSApiClient:
    def __init__(self):
        self.client_id = os.getenv("UPS_CLIENT_ID", "GYLxTKSD7jpnoDnBrGbF2Jk0Uv3MEwPBGSwvBfZQI3DW04nS")
        self.client_secret = os.getenv("UPS_CLIENT_SECRET", "PBJtFdVhZWXfvq6CegMiX3aaUl71cGWkmu2Um5tg4crTkXQUG3hzwehQiGK7EsDg")
        self.account_number = os.getenv("UPS_ACCOUNT_NUMBER", "4a059a")
        self.account_country = os.getenv("UPS_ACCOUNT_COUNTRY", "CA")
        self.env = os.getenv("UPS_ENVIRONMENT", "production") 
        
        if self.env == "production":
            self.base_url = "https://onlinetools.ups.com/api"
        else:
            self.base_url = "https://wwwcie.ups.com/api"
            
        self.token = None

    def get_access_token(self):
        """
        Retrieves OAuth2 access token.
        """
        if self.env == "production":
            url = "https://onlinetools.ups.com/security/v1/oauth/token"
        else:
            url = "https://wwwcie.ups.com/security/v1/oauth/token"
            
        payload = {
            'grant_type': 'client_credentials'
        }
        headers = {
            'Content-Type': 'application/x-www-form-urlencoded'
        }
        
        logging.info(f"Issuing OAuth Token Request to: {url}")
        response = requests.post(url, data=payload, auth=(self.client_id, self.client_secret), headers=headers)
        logging.info(f"OAuth Response {response.status_code}: {response.text}")
        
        if response.status_code == 200:
            self.token = response.json().get('access_token')
            return self.token
        else:
            raise Exception(f"Failed to get access token: {response.text}")

    def create_pickup(self, pickup_data):
        """
        Creates a pickup request using the UPS API.
        """
        tracking_num = pickup_data.get("TrackingNumber", "")
        service_code = pickup_data.get("ServiceCode", "003")
        
        # Map common 1Z service codes to Pickup API service codes (3 digits)
        if tracking_num.startswith("1Z") and len(tracking_num) >= 10:
            label_service = tracking_num[8:10]
            mapping = {
                "01": "001", # Next Day Air
                "02": "002", # 2nd Day Air
                "03": "003", # Ground
                "11": "011", # Standard (Canada Standard Pickup)
                "12": "012", # 3 Day Select
            }
            if label_service in mapping:
                service_code = mapping[label_service]
            else:
                # Default based on country
                country = pickup_data.get("Country", "")
                if country == "CA":
                    service_code = "011"
                elif country == "US":
                    service_code = "003"
                else:
                    service_code = "007"  # UPS Worldwide Express for international
            
        if not self.token:
            self.get_access_token()
            
        url = f"{self.base_url}/pickupcreation/v2403/pickup"
        headers = {
            'Authorization': f'Bearer {self.token}',
            'Content-Type': 'application/json',
            'transId': str(uuid.uuid4()), 
            'transactionSrc': 'testing'
        }
        
        # Basic structure for PickupCreationRequest
        # This is a simplified version, requires careful mapping from implementation details
        request_payload = {
            "PickupCreationRequest": {
                "Request": {
                    "TransactionReference": {
                        "CustomerContext": "StorePickup"
                    }
                },
                "RatePickupIndicator": "N",
                "Shipper": {
                    "Account": {
                        "AccountNumber": self.account_number,
                        "AccountCountryCode": self.account_country
                    }
                },
                "PickupDateInfo": {
                    "CloseTime": pickup_data.get("CloseTime", "1700"),
                    "ReadyTime": pickup_data.get("ReadyTime", "1200"),
                    "PickupDate": pickup_data.get("PickupDate", datetime.datetime.now().strftime("%Y%m%d"))
                },
                "PickupAddress": {
                    "CompanyName": pickup_data.get("CompanyName", "Attention"),
                    "ContactName": pickup_data.get("ContactName", "Warehouse"),
                    "AddressLine": pickup_data.get("Street"),
                    "City": pickup_data.get("City"),
                    "StateProvince": pickup_data.get("State"),
                    "PostalCode": pickup_data.get("Zip"),
                    "CountryCode": pickup_data.get("Country", "US"),
                    "ResidentialIndicator": "N",
                    "PickupPoint": "Lobby",
                    "Phone": {
                        "Number": pickup_data.get("Phone", "0000000000")
                    }
                },
                "AlternateAddressIndicator": "Y",
                "PickupPiece": [
                    {
                        "ServiceCode": self.map_service_code(service_code),
                        "Quantity": "1",
                        "DestinationCountryCode": pickup_data.get("Country", "US"),
                        "ContainerCode": "01"
                    }
                ],
                "TotalWeight": {
                    "Weight": pickup_data.get("Weight", "1.0"),
                    "UnitOfMeasurement": "LBS"
                },
                "OverweightIndicator": "N",
                "TrackingData": [
                    {
                        "TrackingNumber": tracking_num
                    }
                ],
                "PaymentMethod": "04" # 04: Pay by tracking number
            }
        }
        
        # Add Email Notification if provided
        email = pickup_data.get("Email")
        if email:
            request_payload["PickupCreationRequest"]["Notification"] = {
                "ConfirmationEmailAddress": email
            }
        
        logging.info(f"[API Request] create_pickup - URL: {url}\nPayload: {json.dumps(request_payload)}")
        print(f"[API] Creating Pickup for {request_payload['PickupCreationRequest']['PickupDateInfo']['PickupDate']} @ {request_payload['PickupCreationRequest']['PickupDateInfo']['ReadyTime']}-{request_payload['PickupCreationRequest']['PickupDateInfo']['CloseTime']} Local")
        print(f"      Address: {pickup_data.get('Street')} ({pickup_data.get('Country')}) - Tracking: {tracking_num}")
        response = requests.post(url, headers=headers, json=request_payload)
        logging.info(f"[API Response] create_pickup - Status {response.status_code}: {response.text}")
        
        try:
            return response.json()
        except Exception:
            raise Exception(f"HTTP {response.status_code}: {response.text}")

    def create_return_label(self, pickup_data):
        """
        Creates a UPS Return Label shipment and returns the 1Z tracking number.
        Uses ReturnService Code '9' (Print Return Label).
        """
        if not self.token:
            self.get_access_token()
            
        url = self.base_url + "/shipments/v1/ship"
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json",
            "transId": str(uuid.uuid4()),
            "transactionSrc": "OmnitransInc"
        }
        
        # Simplified Return Shipment Payload
        payload = {
            "ShipmentRequest": {
                "Request": {
                    "RequestOption": "nonvalidate",
                    "TransactionReference": {"CustomerContext": "AutoReturnLabel"}
                },
                "Shipment": {
                    "Description": "Return Shipment",
                    "ReturnService": {"Code": "9"}, # 9: Print Return Label
                    "Shipper": {
                        "Name": pickup_data.get("CompanyName", "Warehouse"),
                        "AttentionName": pickup_data.get("ContactName", "Warehouse"),
                        "Phone": {"Number": pickup_data.get("Phone", "5142886664")},
                        "ShipperNumber": self.account_number,
                        "AccountCountryCode": self.account_country, 
                        "Address": {
                            "AddressLine": [pickup_data.get("Street", "")],
                            "City": pickup_data.get("City", ""),
                            "StateProvinceCode": pickup_data.get("State", ""),
                            "PostalCode": pickup_data.get("Zip", ""),
                            "CountryCode": pickup_data.get("Country", "US")
                        }
                    },
                    "ShipTo": {
                        "Name": "Omnitrans Inc",
                        "AttentionName": "Receiving",
                        "Phone": {"Number": "5142886664"},
                        "Address": {
                            "AddressLine": ["9600 rue Meilleur", "Suite 730"],
                            "City": "Montreal",
                            "StateProvinceCode": "QC",
                            "PostalCode": "H2N2E3",
                            "CountryCode": "CA"
                        }
                    },
                    "ShipFrom": {
                        "Name": pickup_data.get("CompanyName", "Warehouse"),
                        "AttentionName": pickup_data.get("ContactName", "Warehouse"),
                        "Phone": {"Number": pickup_data.get("Phone", "5142886664")},
                        "Address": {
                            "AddressLine": [pickup_data.get("Street", "")],
                            "City": pickup_data.get("City", ""),
                            "StateProvinceCode": pickup_data.get("State", ""),
                            "PostalCode": pickup_data.get("Zip", ""),
                            "CountryCode": pickup_data.get("Country", "US")
                        }
                    },
                    "PaymentInformation": {
                        "ShipmentCharge": {
                            "Type": "01",
                            "BillShipper": {"AccountNumber": self.account_number}
                        }
                    },
                    "Service": {"Code": "11"}, # Code 11 is 'UPS Standard' (Required for CA)
                    "Package": {
                        "Description": "Return Package",
                        "Packaging": {"Code": "02"}, # 02: Customer Supplied
                        "PackageWeight": {
                            "UnitOfMeasurement": {"Code": "LBS"},
                            "Weight": "1.0"
                        }
                    }
                },
                "LabelSpecification": {
                    "LabelImageFormat": {"Code": "GIF"},
                    "HTTPUserAgent": "Mozilla/5.0"
                }
            }
        }
        
        logging.info(f"[API Request] create_return_label - URL: {url}\nPayload: {json.dumps(payload)}")
        print(f"[API] Generating Return Label for {pickup_data.get('Street')} ({pickup_data.get('Country')}) -> Montreal (CA) via Service 11")
        response = requests.post(url, headers=headers, json=payload)
        logging.info(f"[API Response] create_return_label - Status {response.status_code}: {response.text}")
        res_json = response.json()
        
        if "ShipmentResponse" in res_json:
            shipment_results = res_json["ShipmentResponse"]["ShipmentResults"]
            # Tracking number is here
            return {
                "status": "success",
                "TrackingNumber": shipment_results["ShipmentIdentificationNumber"]
            }
        return {"status": "error", "message": res_json.get("response", {}).get("errors", [{}])[0].get("message", "Shipping API Error")}

    def map_service_code(self, code):
        """
        Maps 2-digit shipping service codes to 3-digit pickup service codes.
        """
        mapping = {
            "01": "001",
            "02": "002",
            "03": "003",
            "11": "003", # Standard -> Ground Pickup
            "12": "003"
        }
        clean_code = str(code).zfill(2)
        if clean_code in mapping:
            return mapping[clean_code]
        # If already 3 digits or not in mapping, return as is (padded to 3 if needed)
        return str(code).zfill(3)

    def cancel_pickup(self, prn):
        """
        Cancels a scheduled pickup using its PRN.
        """
        if not self.token:
            self.get_access_token()
            
        url = f"{self.base_url}/pickupcreation/v2403/pickup/02"
        headers = {
            'Authorization': f'Bearer {self.token}',
            'transId': f'cancel_{prn}',
            'transactionSrc': 'testing',
            'Prn': prn
        }
        
        logging.info(f"[API Request] cancel_pickup - URL: {url}\nHeaders: {headers}")
        response = requests.delete(url, headers=headers)
        logging.info(f"[API Response] cancel_pickup - Status {response.status_code}: {response.text}")
        
        if response.status_code == 204:
            return {"status": "success", "message": "Pickup cancelled successfully (204 No Content)."}
        try:
            return response.json()
        except:
            return {"status": "error", "message": f"Raw response: {response.text}"}

    def get_pickup_status(self, prn):
        """
        Retrieves real-time status for a scheduled pickup.
        """
        # Always refresh token to avoid stale token 250002 errors
        self.get_access_token()
            
        url = f"{self.base_url}/pickupcreation/v2403/pickup/{prn}"
        headers = {
            'Authorization': f'Bearer {self.token}',
            'transId': f'status_{prn}',
            'transactionSrc': 'testing'
        }
        
        logging.info(f"[API Request] get_pickup_status - URL: {url}")
        response = requests.get(url, headers=headers)
        logging.info(f"[API Response] get_pickup_status - Status {response.status_code}: {response.text}")
        
        if response.status_code == 200:
            return response.json()
        try:
            return response.json()
        except:
            raise Exception(f"HTTP {response.status_code}: {response.text}")

if __name__ == "__main__":
    # This would require credentials in .env to run
    # client = UPSApiClient()
    # print(client.get_access_token())
    pass
