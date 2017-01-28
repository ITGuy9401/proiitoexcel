package eu.arcangelovicedomini.proiiexcel;

public class Point implements Comparable<Point>{
	private Double variablePressureOrTemperature;
	private Double fraction;

	public Point(Double fraction, Double variablePressureOrTemperature) {
		this.variablePressureOrTemperature = variablePressureOrTemperature;
		this.fraction = fraction;
	}

	public Double getVariablePressureOrTemperature() {
		return variablePressureOrTemperature;
	}

	public void setVariablePressureOrTemperature(Double variablePressureOrTemperature) {
		this.variablePressureOrTemperature = variablePressureOrTemperature;
	}

	public Double getFraction() {
		return fraction;
	}

	public void setFraction(Double fraction) {
		this.fraction = fraction;
	}

	public int compareTo(Point o) {
		return variablePressureOrTemperature.compareTo(o.getVariablePressureOrTemperature());
	}
	
}
