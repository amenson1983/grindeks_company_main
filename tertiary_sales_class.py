class Tertiary_sales:
    def __init__(self, basic_list):
        self.penetration = basic_list[6]
        self.year = basic_list[1]
        self.month = basic_list[2]
        self.item = basic_list[0]
        self.brand = basic_list[3]
        self.weight_penetration = basic_list[4]
        self.weight_sro = basic_list[5]
        self.quantity = basic_list[7]
        self.volume_euro = basic_list[8]



    def get_weighted_pen_by_month(self, month):
        list_weight = []
        if self.month == month:
            list_weight.append(self.weight_sro)
        return list_weight




    def __str__(self):
        return f"Year: {self.year}\nMonth: {self.month}\nItem: {self.item}\nBrand: {self.brand}\nWeighted penetration: {self.weight_penetration}\nWeighted SRO: {self.weight_sro}\nQuantity_pcs: {self.quantity}\nAmount_euro: {self.volume_euro}"

    def __repr__(self):
        return f"({self.year},{self.month},{self.item}, {self.brand}, {self.weight_penetration},{self.weight_sro},{self.quantity},{self.volume_euro})"
