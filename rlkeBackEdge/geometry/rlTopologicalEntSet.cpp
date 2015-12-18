#include "stdafx.h"
#include "rlTopologicalEntSet.h"
#include "rlTopologicalEntity.h"
#include "rlVertex2d.h"


rlTopologicalEntSet::rlTopologicalEntSet()
{
	_id_mgr = new rl_id_manager();
}


rlTopologicalEntSet::~rlTopologicalEntSet()
{
	_clear();
	SAFE_DELETE(_id_mgr);
}

void rlTopologicalEntSet::__delete(rlTopologicalEntity *entity)
{
	rlId id = entity->id();
	_ptrs.erase(id);
	SAFE_DELETE(entity);
	_id_mgr->recycle(id);
}

void rlTopologicalEntSet::_clear()
{
	for (std::map<rlId, rlTopologicalEntity*>::iterator it = _ptrs.begin(); it != _ptrs.end(); ++it)
	{
		rlTopologicalEntity *entity = it->second;
		SAFE_DELETE(entity);
	}
	_ptrs.clear();
}

const std::map<rlId, rlTopologicalEntity*> & rlTopologicalEntSet::ptrs() const
{
	return _ptrs;
}

void rlTopologicalEntSet::_insert(rlId id, rlTopologicalEntity *entity)
{
	_ptrs.insert(std::make_pair(id, entity));
}

rl_id_manager *rlTopologicalEntSet::id_mgr() const
{
	return _id_mgr;
}
