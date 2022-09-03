package com.excelcreation.dao;

import org.springframework.data.repository.CrudRepository;
import org.springframework.stereotype.Repository;

import com.excelcreation.entity.Employee;

@Repository
public interface EmployeeDao extends CrudRepository<Employee,Long> {

}
