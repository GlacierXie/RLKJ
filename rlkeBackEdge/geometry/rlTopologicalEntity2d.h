#pragma once
#include "rlTopologicalEntity.h"

class rlTopologicalEntity2d :
	public rlTopologicalEntity
{
public:
	rlSubClassTypeDef(rlTopologicalEntity,emTopologicalEntity2d)
protected:
	rlTopologicalEntity2d();
	~rlTopologicalEntity2d();
};

