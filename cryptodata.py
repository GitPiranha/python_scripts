import pandas as pd


class CryptoData:
    def __init__(self):
        self.customer_id = 0
        self.customer_email = ""
        self.complete_data = pd.DataFrame()
        self.deposit_addresses = pd.DataFrame()
        self.source_addresses = pd.DataFrame()
        self.internal_transaction_ids = pd.DataFrame()
        self.internal_counterparties_ids = pd.DataFrame()
        self.combined_data = pd.DataFrame()

    def read_in_file(self, path):
        self.complete_data = pd.read_excel(path)

    def get_customer_id(self):
        pass

    def get_customer_email(self):
        pass

    def get_deposit_addresses(self):
        pass

    def get_source_addresses(self):
        pass

    def get_internal_transaction_ids(self):
        pass

    def get_internal_counterparty_ids(self):
        pass

    def combine_data(self):
        pass

    def write_data(self):
        pass
