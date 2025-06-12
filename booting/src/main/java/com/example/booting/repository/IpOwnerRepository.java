package com.example.booting.repository;
import com.example.booting.model.IPOwner;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;
import java.util.List;
@Repository
public interface IpOwnerRepository extends JpaRepository<IPOwner, Long>{
List<IPOwner> findByIpAddress(String ipAddress);

}
