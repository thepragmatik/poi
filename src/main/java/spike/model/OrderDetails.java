package spike.model;

public class OrderDetails {

	// POI seems to fall back on Double instead of Long
	private Double orderId;

	private Double quantity;

	private Double unitPrice;

	private String lineItem;

	public Double getOrderId() {
		return orderId;
	}

	public void setOrderId(Double orderId) {
		this.orderId = orderId;
	}

	public Double getQuantity() {
		return quantity;
	}

	public void setQuantity(Double quantity) {
		this.quantity = quantity;
	}

	public Double getUnitPrice() {
		return unitPrice;
	}

	public void setUnitPrice(Double unitPrice) {
		this.unitPrice = unitPrice;
	}

	public String getLineItem() {
		return lineItem;
	}

	public void setLineItem(String lineItem) {
		this.lineItem = lineItem;
	}

}
