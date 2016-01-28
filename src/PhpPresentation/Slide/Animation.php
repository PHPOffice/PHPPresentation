<?php
namespace PhpOffice\PhpPresentation\Slide;
class Animation{
	private $shapeCollection = ARRAY();
	
    public function __construct(){
		
	}
	
	public function addShape($shape){
		$this->shapeCollection[] = $shape;
		
		return $shape;
	}
	
	public function getShapeCollection(){
        return $this->shapeCollection;
    }
	
	
}

?>	
