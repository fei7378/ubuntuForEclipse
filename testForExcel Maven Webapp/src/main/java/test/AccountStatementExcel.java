package test;
public class AccountStatementExcel {
			
	private String name;//商户名称
	private String time;//交易时间
	private String number;//渠道流水号
	private String dmNum;//dm订单号
	private double equipmentNum;//	订单总金额（元）
	private double vipMoney;//会员权益优惠（元）
	private double transactionMoney;//交易手续费（元）
	private double accountMoney;//到账金额（元）
	private String state;//交易状态
	private String by;//支付渠道
	private String scene;//支付场景
	private String terminal;//终端号
	private String playuser;//操作员
	private String otherPs;//备注
	
	
	public AccountStatementExcel(String name, String time, String number, String dmNum, double equipmentNum,
			double vipMoney, double transactionMoney, double accountMoney, String state, String by, String scene,
			String terminal, String playuser, String otherPs) {
		super();
		this.name = name;
		this.time = time;
		this.number = number;
		this.dmNum = dmNum;
		this.equipmentNum = equipmentNum;
		this.vipMoney = vipMoney;
		this.transactionMoney = transactionMoney;
		this.accountMoney = accountMoney;
		this.state = state;
		this.by = by;
		this.scene = scene;
		this.terminal = terminal;
		this.playuser = playuser;
		this.otherPs = otherPs;
	}
	@Override
	public String toString() {
		return "AccountStatementExcel [name=" + name + ", time=" + time + ", number=" + number + ", dmNum=" + dmNum
				+ ", equipmentNum=" + equipmentNum + ", vipMoney=" + vipMoney + ", transactionMoney=" + transactionMoney
				+ ", accountMoney=" + accountMoney + ", state=" + state + ", by=" + by + ", scene=" + scene
				+ ", terminal=" + terminal + ", playuser=" + playuser + ", otherPs=" + otherPs + "]";
	}
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public String getTime() {
		return time;
	}
	public void setTime(String time) {
		this.time = time;
	}
	public String getNumber() {
		return number;
	}
	public void setNumber(String number) {
		this.number = number;
	}
	public String getDmNum() {
		return dmNum;
	}
	public void setDmNum(String dmNum) {
		this.dmNum = dmNum;
	}
	public double getEquipmentNum() {
		return equipmentNum;
	}
	public void setEquipmentNum(double equipmentNum) {
		this.equipmentNum = equipmentNum;
	}
	public double getVipMoney() {
		return vipMoney;
	}
	public void setVipMoney(double vipMoney) {
		this.vipMoney = vipMoney;
	}
	public double getTransactionMoney() {
		return transactionMoney;
	}
	public void setTransactionMoney(double transactionMoney) {
		this.transactionMoney = transactionMoney;
	}
	public double getAccountMoney() {
		return accountMoney;
	}
	public void setAccountMoney(double accountMoney) {
		this.accountMoney = accountMoney;
	}
	public String getState() {
		return state;
	}
	public void setState(String state) {
		this.state = state;
	}
	public String getBy() {
		return by;
	}
	public void setBy(String by) {
		this.by = by;
	}
	public String getScene() {
		return scene;
	}
	public void setScene(String scene) {
		this.scene = scene;
	}
	public String getPlayuser() {
		return playuser;
	}
	public void setPlayuser(String playuser) {
		this.playuser = playuser;
	}
	public String getTerminal() {
		return terminal;
	}
	public void setTerminal(String terminal) {
		this.terminal = terminal;
	}
	public String getOtherPs() {
		return otherPs;
	}
	public void setOtherPs(String otherPs) {
		this.otherPs = otherPs;
	}
	
	
	

}
