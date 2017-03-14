package accounting.gulfmark.models;

public class VolEntry {
	private String month;
	private String year;
	private String volNumber;
	private String volName;
	private String county;
	private String state;
	private Double unit;
	private Double volume;
	private Double net;
	
	
	public VolEntry() {
		super();
	}
	public VolEntry(String month, String year, String volNumber, String volName, String county, String state, Double unit, Double volumn,
			Double net) {
		super();
		this.month = month;
		this.year = year;
		this.volNumber = volNumber;
		this.volName = volName;
		this.county = county;
		this.state = state;
		this.unit = unit;
		this.volume = volumn;
		this.net = net;
	}
	
	
	public String getMonth() {
		return month;
	}
	public void setMonth(String month) {
		this.month = month;
	}
	public String getYear() {
		return year;
	}
	public void setYear(String year) {
		this.year = year;
	}
	public String getVolNumber() {
		return volNumber;
	}
	public void setVolNumber(String volNumber) {
		this.volNumber = volNumber;
	}
	public String getVolName() {
		return volName;
	}
	public void setVolName(String volName) {
		this.volName = volName;
	}
	public String getCounty() {
		return county;
	}
	public void setCounty(String county) {
		this.county = county;
	}
	public String getState() {
		return state;
	}
	public void setState(String state) {
		this.state = state;
	}
	public Double getUnit() {
		return unit;
	}
	public void setUnit(Double unit) {
		this.unit = unit;
	}
	public Double getVolume() {
		return volume;
	}
	public void setVolume(Double volumn) {
		this.volume = volumn;
	}
	public Double getNet() {
		return net;
	}
	public void setNet(Double net) {
		this.net = net;
	}
}
