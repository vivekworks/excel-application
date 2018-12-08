package model.java;

public class EODExcel implements Comparable<EODExcel> {
    private String dateCreated, srNo, description, srType, customer, version, country, mosStatus, notes, colorOrder;
    private boolean duplicateCheck;

    public EODExcel() {
        super();
    }

    public void setDateCreated(String dateCreated) {
        if (dateCreated != null && !(dateCreated.equalsIgnoreCase("")))
            this.dateCreated = dateCreated.trim();
    }

    public String getDateCreated() {
        return dateCreated;
    }

    public void setSrNo(String srNo) {
        if (srNo != null && !(srNo.equalsIgnoreCase("")))
            this.srNo = srNo.trim();
    }

    public String getSrNo() {
        return srNo;
    }

    public void setDescription(String description) {
        if (description != null && !(description.equalsIgnoreCase("")))
            this.description = description.trim();
    }

    public String getDescription() {
        return description;
    }

    public void setSrType(String srType) {
        if (srType != null && !(srType.equalsIgnoreCase("")))
            this.srType = srType.replace("-", " ").toUpperCase().trim();
    }

    public String getSrType() {
        return srType;
    }

    public void setCustomer(String customer) {
        if (customer != null && !(customer.equalsIgnoreCase("")))
            this.customer = customer.trim();
    }

    public String getCustomer() {
        return customer;
    }

    public void setVersion(String version) {
        if (version != null && !(version.equalsIgnoreCase("")))
            this.version = version.trim();
    }

    public String getVersion() {
        return version;
    }

    public void setCountry(String country) {
        if (country != null && !(country.equalsIgnoreCase("")))
            this.country = country.trim();
    }

    public String getCountry() {
        return country;
    }

    public void setMosStatus(String mosStatus) {
        if (mosStatus != null && !(mosStatus.equalsIgnoreCase("")))
            this.mosStatus = mosStatus.trim();
    }

    public String getMosStatus() {
        return mosStatus;
    }

    public void setNotes(String notes) {
        if (notes != null && !(notes.equalsIgnoreCase("")))
            this.notes = notes.trim();
    }

    public String getNotes() {
        return notes;
    }

    public int compareTo(EODExcel eodEx) {
        return (this.getSrNo() + this.getColorOrder()).compareTo(eodEx.getSrNo() + eodEx.getColorOrder());
    }

    public void setDuplicateCheck(boolean duplicateCheck) {
        this.duplicateCheck = duplicateCheck;
    }

    public boolean isDuplicateCheck() {
        return duplicateCheck;
    }

    public void setColorOrder(String colorOrder) {
        if (colorOrder != null && !(colorOrder.equalsIgnoreCase("")))
            this.colorOrder = colorOrder.trim();
    }

    public String getColorOrder() {
        return colorOrder;
    }
}
