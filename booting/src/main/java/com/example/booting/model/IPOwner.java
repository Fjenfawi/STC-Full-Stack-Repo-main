package com.example.booting.model;
import org.springframework.stereotype.Indexed;

import jakarta.persistence.Entity;
import jakarta.persistence.Table;
import jakarta.persistence.Id;
import jakarta.persistence.GeneratedValue;
import jakarta.persistence.GenerationType;


@Entity
@Table(name= "ip_owner")
public class IPOwner {
@Id
@GeneratedValue(strategy = GenerationType.IDENTITY)
private Long id;
private String ipAddress;
private String ownerName;


//constructor
public IPOwner(){}

public IPOwner(String ipAddress, String ownerName){
    this.ipAddress = ipAddress;
    this.ownerName = ownerName;

}

//getters & setters

public Long getId(){
    return id;
}
public String getIpAddress(){
    return ipAddress;

}
public String getOwnerName(){
    return ownerName;

}
public void setIpAddress(String ipAddress){
this.ipAddress = ipAddress;
}
public void setOwnerName(String ownerName){
this.ownerName = ownerName;
}

//toString for debugging

@Override
public String toString(){
    return "IpOwner{" +
               "id=" + id +
               ", ipAddress='" + ipAddress + '\'' +
               ", ownerName='" + ownerName + '\'' +
               '}';
}

}
