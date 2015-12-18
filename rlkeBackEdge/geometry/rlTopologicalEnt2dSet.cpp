#include "stdafx.h"
#include "rlTopologicalEnt2dSet.h"
#include "rlEdge2d.h"
#include "rlVertex2d.h"


rlTopologicalEnt2dSet::rlTopologicalEnt2dSet()
{
}


rlTopologicalEnt2dSet::~rlTopologicalEnt2dSet()
{
}

void rlTopologicalEnt2dSet::_delete(rlTopologicalEntity2d *entity)
{
	rlTopologicalEntSet::__delete(entity);
}

std::map<rlId, rlEdge2d *> rlTopologicalEnt2dSet::edges() const
{
	std::map<rlId, rlEdge2d *> out;

	const std::map<rlId, rlTopologicalEntity *> &_ptrs = ptrs();
	std::map<rlId, rlTopologicalEntity *>::const_iterator it = _ptrs.begin();
	for (; it != _ptrs.end(); ++it)
	{
		rlTopologicalEntity *entity = it->second;
		if (entity != NULL && entity->type() == emEdge2d)
			out.insert(std::make_pair(it->first, dynamic_cast<rlEdge2d *>(it->second)));
	}

	return out;
}

std::map<rlId, rlVertex2d *> rlTopologicalEnt2dSet::vertices() const
{
	std::map<rlId, rlVertex2d *> out;

	const std::map<rlId, rlTopologicalEntity *> &_ptrs = ptrs();
	std::map<rlId, rlTopologicalEntity *>::const_iterator it = _ptrs.begin();
	for (; it != _ptrs.end(); ++it)
	{
		rlTopologicalEntity *entity = it->second;
		if (entity != NULL && entity->type() == emVertex2d)
			out.insert(std::make_pair(it->first, dynamic_cast<rlVertex2d *>(it->second)));
	}

	return out;
}
