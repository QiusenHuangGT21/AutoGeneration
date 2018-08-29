import javafx.beans.property.SimpleFloatProperty;
import javafx.beans.property.SimpleStringProperty;

public class Car {
    private String date;
    private String brandName;
    private Float weight;
    private SimpleFloatProperty price;
    private SimpleStringProperty priceString;
    private SimpleFloatProperty total = new SimpleFloatProperty(0);
    private String carNumber;
    private boolean type = false;
    private String site;

    public Car(){

    }

    public Car(String line){
        String[] dataField = line.split(" ");
        this.date = dataField[0];
        this.brandName = dataField[2];
        this.weight = Float.parseFloat(dataField[3]);
        this.carNumber = dataField[1];
        this.setPrice(Float.parseFloat(dataField[4]));
        this.total = new SimpleFloatProperty(Float.parseFloat(dataField[5]));
        this.site = dataField[6];
        if (brandName.contains("散")) {
            this.type = true;
        }
    }

    public Car(String date, String brandName, Float weight, String carNumber, String site) {
        this.date = date;
        this.brandName = brandName;
        this.weight = weight;
        this.carNumber = carNumber;
        this.site = site;
        if (brandName.contains("散")) {
            this.type = true;
        }
    }

    public String getDate() {
        return date;
    }

    public void setDate(String date) {
        this.date = date;
    }

    public SimpleStringProperty dateProperty() {
        return new SimpleStringProperty(this.date);
    }

    public String getBrandName() {
        return brandName;
    }

    public void setBrandName(String brandName) {
        this.brandName = brandName;
    }

    public SimpleStringProperty brandNameProperty() {
        return new SimpleStringProperty(this.brandName);
    }

    public Float getWeight() {
        return weight;
    }

    public void setWeight(Float weight) {
        this.weight = weight;
    }

    public SimpleFloatProperty weightProperty() {
        return new SimpleFloatProperty(this.weight);
    }

    public SimpleFloatProperty getPrice() {
        return price;
    }

    public void setPrice(Float price) {
        this.price = new SimpleFloatProperty(price);
        this.total.bind(this.price.multiply(this.getWeight()));
        this.priceString = new SimpleStringProperty(price.toString());
    }

    public SimpleFloatProperty getTotal() {
        return total;
    }

    public SimpleFloatProperty totalProperty() {
        return this.getTotal();
    }

    public String getCarNumber() {
        return carNumber;
    }

    public void setCarNumber(String carNumber) {
        this.carNumber = carNumber;
    }

    public SimpleStringProperty carNumberProperty() {
        return new SimpleStringProperty(this.getCarNumber());
    }

    public String getSite() {
        return site;
    }

    public void setSite(String site) {
        this.site = site;
    }

    public SimpleStringProperty siteProperty() {
        return new SimpleStringProperty(this.getSite());
    }

    public SimpleStringProperty getPriceString() {
        return this.priceString;
    }

    public void setPriceString(String priceString) {
        this.priceString = new SimpleStringProperty(priceString);
        this.setPrice(new Float(priceString));
    }

    public SimpleStringProperty priceStringProperty() {
        return this.getPriceString();
    }
    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;

        Car car = (Car) o;

        if (!getDate().equals(car.getDate())) return false;
        if (!getBrandName().equals(car.getBrandName())) return false;
        if (!getWeight().equals(car.getWeight())) return false;
        if (getPrice() != null ? !getPrice().equals(car.getPrice()) : car.getPrice() != null) return false;
        if (getTotal() != null ? !getTotal().equals(car.getTotal()) : car.getTotal() != null) return false;
        if (!getCarNumber().equals(car.getCarNumber())) return false;
        if (!getSite().equals(car.getSite())) return false;
        if (isType() != car.isType()) return false;
        return true;
    }

    @Override
    public int hashCode() {
        int result = getDate().hashCode();
        result = 31 * result + getBrandName().hashCode();
        result = 31 * result + getWeight().hashCode();
        result = 31 * result + (getPrice() != null ? getPrice().hashCode() : 0);
        result = 31 * result + (getTotal() != null ? getTotal().hashCode() : 0);
        result = 31 * result + getCarNumber().hashCode();
        result = 31 * result + getSite().hashCode();
        return result;
    }

    @Override
    public String toString() {
        return "" + date + " " +
                carNumber + ' ' +
                brandName + ' ' +
                weight + ' ' +
                (price == null ? 0 : price.get()) + ' ' +
                (total == null ? 0 : total.get()) + ' ' +
                site + ' ' + "\n"
                ;
    }

    public boolean isType() {
        return type;
    }

    public void setType(boolean type) {
        this.type = type;
    }
}
