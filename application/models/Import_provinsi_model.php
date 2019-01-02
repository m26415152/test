<?php
class Import_provinsi_model extends CI_Model{
	 function select(){
	 	$this->db->order_by('id', 'DESC');
	 	$query = $this->db->get('provinsi');
	 	return $query;
	 }

	 function insert(){
	 	$this->db->insert_batch('provinsi', $data);	
	 }
	 
} 