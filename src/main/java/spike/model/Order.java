package spike.model;

import java.util.Date;

public class Order {

	// POI seems to fall back on Double instead of Long
	private Double id;

	private String name;

	private Date orderDate;

	public Double getId() {
		return id;
	}

	public void setId(Double id) {
		this.id = id;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public Date getOrderDate() {
		return orderDate;
	}

	public void setOrderDate(Date orderDate) {
		this.orderDate = orderDate;
	}

}
