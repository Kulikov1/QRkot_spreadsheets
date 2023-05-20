from sqlalchemy import select, func
from sqlalchemy.ext.asyncio import AsyncSession

from app.crud.base import CRUDBase
from app.models import CharityProject


class CRUDCharityProject(CRUDBase):
    async def get_projects_by_completion_rate(
        self,
        session: AsyncSession,
    ):
        charity_projects = await session.execute(
            select(CharityProject).where(
                CharityProject.fully_invested
            ).order_by(
                func.julianday(CharityProject.close_date) -
                func.julianday(CharityProject.create_date)
            )
        )
        charity_projects = charity_projects.scalars().all()
        return charity_projects


charity_project_crud = CRUDCharityProject(CharityProject)
