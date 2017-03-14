package accounting.shell.models;

public class Settlement {
	private String period;
	private String buySellFlag;
	private String status;
	private String date;
	private String smartNo;
	private String unit;
	private String currency;
	private String cashFlowType;
	private String location;
	private String leaseNo;
	private String leaseName;
	private String product;
	private String pipeline;
	private String agreement;

	private int contractNo;
	private int commitment;
	private int dealTrackNo;
	
	private double volume;
	private double price;
	private double settleAmount;
	
	public Settlement() {
		super();
	}

	public String getPeriod() {
		return period;
	}

	public void setPeriod(String period) {
		this.period = period;
	}
	public String getBuySellFlag() {
		return buySellFlag;
	}

	public void setBuySellFlag(String buySellFlag) {
		this.buySellFlag = buySellFlag;
	}
	
	public String getStatus() {
		return status;
	}

	public void setStatus(String status) {
		this.status = status;
	}

	public String getDate() {
		return date;
	}

	public void setDate(String date) {
		this.date = date;
	}

	public int getContractNo() {
		return contractNo;
	}

	public void setContractNo(int contractNo) {
		this.contractNo = contractNo;
	}

	public String getSmartNo() {
		return smartNo;
	}

	public void setSmartNo(String smartNo) {
		this.smartNo = smartNo;
	}

	public int getDealTrackNo() {
		return dealTrackNo;
	}

	public void setDealTrackNo(int dealTrackNo) {
		this.dealTrackNo = dealTrackNo;
	}

	public String getUnit() {
		return unit;
	}

	public void setUnit(String unit) {
		this.unit = unit;
	}

	public String getCurrency() {
		return currency;
	}

	public void setCurrency(String currency) {
		this.currency = currency;
	}

	public String getProduct() {
		return product;
	}

	public void setProduct(String product) {
		this.product = product;
	}

	public String getPipeline() {
		return pipeline;
	}

	public void setPipeline(String pipeline) {
		this.pipeline = pipeline;
	}

	public String getAgreement() {
		return agreement;
	}

	public void setAgreement(String agreement) {
		this.agreement = agreement;
	}

	public String getCashFlowType() {
		return cashFlowType;
	}

	public void setCashFlowType(String cashFlowType) {
		this.cashFlowType = cashFlowType;
	}

	public String getLocation() {
		return location;
	}

	public void setLocation(String location) {
		this.location = location;
	}

	public String getLeaseNo() {
		return leaseNo;
	}

	public void setLeaseNo(String leaseNo) {
		this.leaseNo = leaseNo;
	}

	public String getLeaseName() {
		return leaseName;
	}

	public void setLeaseName(String leaseName) {
		this.leaseName = leaseName;
	}

	public int getCommitment() {
		return commitment;
	}

	public void setCommitment(int commitment) {
		this.commitment = commitment;
	}

	public double getVolume() {
		return volume;
	}

	public void setVolume(double volume) {
		this.volume = volume;
	}

	public double getPrice() {
		return price;
	}

	public void setPrice(double price) {
		this.price = price;
	}

	public double getSettleAmount() {
		return settleAmount;
	}

	public void setSettleAmount(double settleAmount) {
		this.settleAmount = settleAmount;
	}
	
}
