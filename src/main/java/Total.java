import javafx.beans.property.SimpleFloatProperty;
import javafx.beans.property.SimpleStringProperty;

import java.math.BigDecimal;

public class Total {
    private SimpleFloatProperty weightTotal = new SimpleFloatProperty();
    private SimpleFloatProperty priceTotal = new SimpleFloatProperty();
    private SimpleStringProperty rmbTotal = new SimpleStringProperty();
    private String date;

    public Total() {
        this.date = "";
        this.weightTotal.set(0);
        this.priceTotal.set(0);
        this.rmbTotal.set(CNY.number2CNMontrayUnit(new BigDecimal(0)));

    }

    public Total(String date) {
        this.weightTotal.set(0);
        this.priceTotal.set(0);
        this.rmbTotal.set(CNY.number2CNMontrayUnit(new BigDecimal(0)));
        this.date = date;
    }

    public void increment(Car c) {
        this.weightTotal.add(c.getWeight());
        this.priceTotal.add(c.getPrice().floatValue() * c.getWeight());
    }

    public void setRmbTotal() {
        this.rmbTotal = new SimpleStringProperty(CNY.number2CNMontrayUnit(new BigDecimal(priceTotal.get())));
    }

    public String getDate() {
        return this.date;
    }

    public SimpleFloatProperty weightTotalProperty() {
        return this.weightTotal;
    }

    public SimpleFloatProperty priceTotalProperty() {
        return this.priceTotal;
    }

    public SimpleStringProperty rmbTotalProperty() {
        return this.rmbTotal;
    }

    @Override
    public String toString() {
        return this.date + this.rmbTotal.getName() + " " + this.priceTotal.getName()
                + " " + this.weightTotal.getName() + this.rmbTotal + "\n";
    }
}