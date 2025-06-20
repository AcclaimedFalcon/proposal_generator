
def calculate_total_after_gst(base_price, gst):
    return base_price + gst

def calculate_net_payable(total_after_gst, subsidy):
    return total_after_gst - subsidy
