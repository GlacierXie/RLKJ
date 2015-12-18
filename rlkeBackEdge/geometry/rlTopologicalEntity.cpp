#include "stdafx.h"
#include "rlTopologicalEntity.h"


rlTopologicalEntity::rlTopologicalEntity()
{
}


rlTopologicalEntity::~rlTopologicalEntity()
{
}

rlId rlTopologicalEntity::id() const
{
	return _id;
}

void rlTopologicalEntity::_set(rlId id)
{
	_id = id;
}
