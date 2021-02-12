class Tertiary_sales:
    def __init__(self, basic_list):
        self.penetration = basic_list[6]
        self.year = basic_list[1]
        self.month = basic_list[2]
        self.item = basic_list[0]
        self.brand = basic_list[3]
        self.weight_penetration = basic_list[4]
        self.sro = basic_list[5]
        self.quantity = basic_list[7]
        self.volume_euro = basic_list[8]
        self.weighted_sro = basic_list[9]


    def __str__(self):
        return f"Year: {self.year}\nMonth: {self.month}\nItem: {self.item}\nBrand: {self.brand}\nWeighted penetration: {self.weight_penetration}\nWeighted SRO: {self.sro}\nQuantity_pcs: {self.quantity}\nAmount_euro: {self.volume_euro}"

    def __repr__(self):
        return f"({self.year},{self.month},{self.item}, {self.brand}, {self.weight_penetration},{self.sro},{self.quantity},{self.volume_euro})"
