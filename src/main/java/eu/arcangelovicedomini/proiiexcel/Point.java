package eu.arcangelovicedomini.proiiexcel;

public class Point {
	private Double firstColumn;
	private Double secondColumn;

	public Point(Double firstColumn, Double secondColumn) {
		this.firstColumn = firstColumn;
		this.secondColumn = secondColumn;
	}

	public Double getFirstColumn() {
		return firstColumn;
	}

	public void setFirstColumn(Double firstColumn) {
		this.firstColumn = firstColumn;
	}

	public Double getSecondColumn() {
		return secondColumn;
	}

	public void setSecondColumn(Double secondColumn) {
		this.secondColumn = secondColumn;
	}

}
