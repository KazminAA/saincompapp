import java.util.Objects;

public class ParsedRow {
    private String operation;
    private String sain;
    private String name;
    private Integer number;
    private String notes;

    public String getSain() {
        return sain;
    }

    public void setSain(String sain) {
        this.sain = sain;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Integer getNumber() {
        return number;
    }

    public void setNumber(Integer number) {
        this.number = number;
    }

    public String getOperation() {
        return operation;
    }

    public void setOperation(String operation) {
        this.operation = operation;
    }

    public String getNotes() {
        return notes;
    }

    public void setNotes(String notes) {
        this.notes = notes;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        ParsedRow parsedRow = (ParsedRow) o;
        return Objects.equals(sain, parsedRow.sain) &&
                Objects.equals(name, parsedRow.name);
    }

    @Override
    public int hashCode() {

        return Objects.hash(sain, name);
    }

    @Override
    public String toString() {
        return "ParsedRow{" +
                "sain='" + sain + '\'' +
                ", name='" + name + '\'' +
                ", number=" + number +
                '}';
    }
}
